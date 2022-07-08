'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ELIGIBILITY SUMMARY.vbs"
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
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("07/05/2022", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


function ensure_variable_is_a_number(variable)
	variable = trim(variable)
	If variable = "" Then variable = 0
	variable = variable * 1
	variable = variable + 0
end function

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result, approval_found)
	Call write_value_and_transmit("99", cmd_row, cmd_col)
	approval_found = True

	row = 7
	Do
		EMReadScreen elig_version, 2, row, 22
		EMReadScreen elig_date, 8, row, 26
		EMReadScreen elig_result, 10, row, 37
		EMReadScreen approval_status, 10, row, 50

		elig_version = trim(elig_version)
		elig_result = trim(elig_result)
		approval_status = trim(approval_status)

		If approval_status = "APPROVED" Then Exit Do

		row = row + 1
	Loop until approval_status = ""

	Call clear_line_of_text(18, 54)
	If approval_status = "" Then
		approval_found = false
		PF3
	Else
		Call write_value_and_transmit(elig_version, 18, 54)
		version_number = "0" & elig_version
		version_date = elig_date
		version_result = elig_result
	End If
end function

function determine_130_percent_of_FPG(footer_month, footer_year, hh_size, fpg_130_percent)

	month_to_review = footer_month & "/1/" & footer_year
	month_to_review = DateAdd("d", 0, month_to_review)

	If IsNumeric(hh_size) = True Then
		hh_size = hh_size*1

		If DateDiff("d", #10/1/2021#, month_to_review) >= 0 Then
			If hh_size = 1 Then fpg_130_percent = 1396
			If hh_size = 2 Then fpg_130_percent = 1888
			If hh_size = 3 Then fpg_130_percent = 2379
			If hh_size = 4 Then fpg_130_percent = 2871
			If hh_size = 5 Then fpg_130_percent = 3363
			If hh_size = 6 Then fpg_130_percent = 3855
			If hh_size = 7 Then fpg_130_percent = 4347
			If hh_size = 8 Then fpg_130_percent = 4839

			If hh_size > 8 Then fpg_130_percent = 4839 + (492 * (hh_size-8))
		End If
	End If

end function

function snap_elig_dialog()

	BeginDialog Dialog1, 0, 0, 555, 385, "SNAP Approval Packages"
	  GroupBox 460, 10, 85, 165, "SNAP Approvals"

	  Text 10, 355, 175, 10, "Confirm you have reviewed the budget for accuracy:"
	  DropListBox 185, 350, 155, 45, "Indicate if the Budget is Accurate"+chr(9)+"Yes - budget is Accurate"+chr(9)+"No - I need to complete a new Approval", SNAP_UNIQUE_APPROVALS(confirm_budget_selection, approval_selected)

	  If SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, approval_selected) = True Then
	  	GroupBox 5, 10, 285, 105, "Approval Detail"
	  	Text 15, 20, 135, 10, "Total Gross Income . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_gross_inc
	  	Text 15, 30, 135, 10, "Total Deductions . . . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_deduct & "  ( - )"
	  	Text 15, 40, 135, 10, "Net Income . . . . . . . . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_net_inc
	  	Text 15, 50, 135, 10, "Shelter Expense . . . . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_shel_expenses & "  ( - )"
	  	Text 15, 60, 135, 10, "Net Adjusted Income . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_net_adj_inc

	  	Text 15, 75, 135, 10, "Thrifty Food Plan . . . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_thrifty_food_plan
	  	Text 15, 85, 135, 10, "30% of Net Adj Income . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_bug_30_percent_net_adj_inc & "  ( - )"

	  	Text 40, 100, 105, 10, "Entitlement . . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_benefit_monthly_fs_allot
	  	' Text 15, 95, 130, 25, "Monthly SNAP Allotment calculated by subtracting 30% of the adjusted net income from the Thrifty Food Plan"
	  	Text 165, 20, 120, 10, "Months in Approval: " & display_detail
	  	Text 205, 30, 80, 10, " Result:   " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result

	  	Text 165, 40, 120, 10, " Benefit Entitlement:   $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_benefit_monthly_fs_allot
	  	Text 165, 60, 115, 10, "Max Gross Inc . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_gross_inc
	  	Text 170, 70, 110, 10, "Gross Income Test . . . " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_prosp_gross_inc_test
	  	Text 165, 80, 115, 10, "Max Net Inc . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_net_adj_inc
	  	Text 170, 90, 110, 10, "Net Income Test . . . . . " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_prosp_net_inc_test

	  	GroupBox 300, 10, 150, 80, "Total Deductions Calculation"
	  	Text 325, 35, 100, 10, " Standard . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard
	  	Text 320, 45, 100, 10, " Earned Inc . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned
	  	Text 330, 55, 100, 10, "Medical . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical
	  	Text 305, 65, 130, 10, "Dependent Care . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care
	  	Text 310, 75, 130, 10, " Child Support . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses

	  	GroupBox 300, 95, 150, 80, "Allowable Shelter Cost Calculation"
	  	Text 305, 110, 145, 10, "Total Shelter Costs . . . . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_shel_total
	  	Text 305, 120, 145, 10, "Half of Net Income . . . . . . . . .$ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_50_perc_net_inc & "  ( - )"
	  	Text 305, 130, 145, 10, "Adjusted Shelter Costs . . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_adj_shel_costs
	  	Text 305, 140, 90, 20, "This case has a maximum shelter cost of $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_allow_shel
	  	Text 305, 160, 145, 10, "Allowed Shelter Expense . . . . $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_shel_expenses

	  Else

	  	GroupBox 5, 10, 450, 90, "Approval Detail"

	  	Text 15, 20, 120, 10, "Months in Approval: " & display_detail
	  	Text 55, 30, 80, 10, " Result:   " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result

	  	Text 15, 45, 100, 10, "APPL Withdrawn:    " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_appl_withdrawn_test
	  	Text 15, 55, 100, 10, "Applicant Elig:         " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_applct_elig_test
	  	Text 15, 65, 100, 10, "Commodity:             " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_comdty_test
	  	Text 15, 75, 100, 10, "Disqualification:      " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_disq_test
	  	Text 15, 85, 100, 10, "Duplicate Assist:     " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_dupl_assist_test

	  	Text 125, 45, 100, 10, "Eligible Person:       " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_eligible_person_test
	  	Text 125, 55, 100, 10, "Fail Cooperation:     " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_fail_coop_test
	  	Text 125, 65, 100, 10, "Fail to File:               " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_fail_file_test
	  	Text 125, 75, 100, 10, "Prosp Gross Inc:     " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_prosp_gross_inc_test
	  	Text 125, 85, 100, 10, "Prosp Net Inc:         " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_prosp_net_inc_test

	  	Text 235, 45, 100, 10, "Recertification:     " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_recert_test
	  	Text 235, 55, 100, 10, "Residence:           " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_residence_test
	  	Text 235, 65, 100, 10, "Resource:             " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_resource_test
	  	Text 235, 75, 100, 10, "Retro Gross Inc:    " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_retro_gross_inc_test
	  	Text 235, 85, 100, 10, "Retro Net Inc:        " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_retro_net_inc_test

	  	Text 345, 45, 100, 10, "Strike:                    " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_strike_test
	  	Text 345, 55, 100, 10, "Xfer Asset/Inc:      " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_xfer_resource_inc_test
	  	Text 345, 65, 100, 10, "Verification:            " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test
	  	Text 345, 75, 100, 10, "Voluntary Quit:       " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_voltry_quit_test
	  	Text 345, 85, 100, 10, "Work Registration: " & SNAP_ELIG_APPROVALS(elig_ind).snap_case_work_reg_test

	  	GroupBox 5, 100, 450, 60, "Ineligible Details"
	  	If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test = "FAILED" then
			Text 15, 115, 165, 10, "What is the date the verification request was sent? "
			Editbox 180, 110, 50, 15, SNAP_UNIQUE_APPROVALS(verif_reqquest_date, approval_selected)
			Text 235, 115, 150, 10, "(due date is 10 days from this request date)"
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_PACT = "FAILED" then
				Text 15, 135, 120, 10, "List PACT reason(s) for ineligibility: "
				Editbox 130, 130, 310, 15, SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected)
				Text 130, 145, 300, 10, "Phrase this for residents as this detail will be added to the WCOM."
			End if
		Else
			Text 15, 120, 300, 10, "This case is ineligible becaues it hasn't met the requirements for SNAP Eligibility. The case tests above show what requirements have not been met."
	    End if

	  End If

	  ' EditBox 600, 400, 50, 10, empty_editbox
	  ButtonGroup ButtonPressed
	    If SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, approval_selected) = False Then PushButton 390, 115, 50, 10, "View ELIG", nav_stat_elig_btn
		If SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, approval_selected) = True Then PushButton 165, 100, 50, 10, "View ELIG", nav_stat_elig_btn
		If snap_approval_is_incorrect = True Then
			PushButton 440, 365, 110, 15, "Cancel Approval Noting", app_incorrect_btn
		ElseIf approval_selected = UBound(SNAP_UNIQUE_APPROVALS, 2) or snap_approval_is_incorrect = True Then
			PushButton 440, 365, 110, 15, "Approvals Confirmed", app_confirmed_btn
		Else
			PushButton 440, 365, 110, 15, "Next Approval", next_approval_btn
		End If
		If SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, approval_selected) = True Then PushButton 360, 20, 85, 10, "Deductions Detail", deductions_detail_btn
		PushButton 200, 160, 70, 10, "HH COMP Detail", hh_comp_detail
		PushButton 360, 190, 85, 10, "Shelter Expense Detail", shel_exp_detail_btn
		y_pos = 25
		display_detail = ""
		for each_app = 0 to UBound(SNAP_UNIQUE_APPROVALS, 2)
			If SNAP_UNIQUE_APPROVALS(last_mo_const, each_app) = "" Then
				month_display = SNAP_UNIQUE_APPROVALS(first_mo_const, each_app)
			ElseIF SNAP_UNIQUE_APPROVALS(last_mo_const, each_app) = CM_plus_1_mo & "/" & CM_plus_1_yr Then
				month_display = SNAP_UNIQUE_APPROVALS(first_mo_const, each_app) & " - Ongoing"
			Else
				month_display = SNAP_UNIQUE_APPROVALS(first_mo_const, each_app) & " - " & SNAP_UNIQUE_APPROVALS(last_mo_const, each_app)
			End if
			If each_app = approval_selected Then display_detail = month_display
			If each_app = approval_selected Then
				Text 470, y_pos+2, 75, 13, month_display
			Else
				PushButton 465, y_pos, 75, 13, month_display, SNAP_UNIQUE_APPROVALS(btn_one, each_app)
			End If
			y_pos = y_pos + 15
		next
		PushButton 465, 150, 75, 20, "About Approval Pkgs", unique_approval_explain_btn

	  If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "ELIGIBLE" Then
		  GroupBox 5, 120, 285, 35, "SNAP Benefits Issued to Resident in the Approval Package"
		  app_y_pos = 132
		  app_x_pos = 10
		  For approval = 0 to UBound(SNAP_ELIG_APPROVALS)
		  	If InStr(SNAP_UNIQUE_APPROVALS(months_in_approval, approval_selected), SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year) <> 0 Then
				Text app_x_pos, app_y_pos, 85, 10, SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year & " - $ " & SNAP_ELIG_APPROVALS(approval).snap_benefit_amt
				app_y_pos = app_y_pos + 10
				If app_y_pos = 152 Then
					app_y_pos = 132
					app_x_pos = app_x_pos + 90
				End If
			End If
		  Next
	  End If

	  GroupBox 5, 160, 285, 70, "Household Composition"
	  Text 20, 170, 285, 10, "Members in Assistance Unit:  " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_numb_in_assist_unit & " - Adult: " & SNAP_ELIG_APPROVALS(elig_ind).adults_recv_snap & ", Children: " & SNAP_ELIG_APPROVALS(elig_ind).children_recv_snap
	  Text 35, 180, 245, 20, "Eligible Members:  " & SNAP_ELIG_APPROVALS(elig_ind).elig_membs_list
	  Text 30, 200, 245, 20, "Ineligible Members:  " & SNAP_ELIG_APPROVALS(elig_ind).inelig_membs_list

	  ' GroupBox 300, 90, 150, 80, "Allowable Shelter Cost Calculation"
	  GroupBox 300, 180, 240, 50, "Expenses"
	  Text 315, 205, 200, 10, "Utilities Expense:   $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_utilities_exp_total & "  -  " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_utilities_list
	  Text 310, 215, 125, 10, "Housing Expense:  $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_housing_exp_total

	  GroupBox 5, 235, 540, income_box_len, "Income"
	  Text 10, 245, 155, 10, "Total GROSS EARNED Income:   $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc
	  Text 300, 245, 155, 10, "Total GROSS UNEARNED Income:   $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_unea_inc
	  y_pos = 260
	  y_pos_2 = 260
	  For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	  	If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_one_job_counted(each_memb) = True Then
			' Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb)
			Text 15, y_pos, 215, 10, "$ " & STAT_INFORMATION(month_ind).stat_jobs_one_snap_pic_prosp_monthly_inc(each_memb) & " - Monthly Income   --   Memb " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb)
  		    If STAT_INFORMATION(month_ind).stat_jobs_one_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 40, y_pos+10, 200, 10, "Paid " & STAT_INFORMATION(month_ind).stat_jobs_one_snap_pic_pay_freq(each_memb) & "   --   $ " & STAT_INFORMATION(month_ind).stat_jobs_one_snap_pic_ave_inc_per_pay(each_memb) & " average inc/pay date"
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_two_job_counted(each_memb) = True Then
			' Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_jobs_two_snap_pic_prosp_monthly_inc(each_memb)
			Text 15, y_pos, 215, 10, "$ " & STAT_INFORMATION(month_ind).stat_jobs_two_snap_pic_prosp_monthly_inc(each_memb) & " - Monthly Income   --   Memb " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb)
			If STAT_INFORMATION(month_ind).stat_jobs_two_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 40, y_pos+10, 200, 10, "Paid " & STAT_INFORMATION(month_ind).stat_jobs_two_snap_pic_pay_freq(each_memb) & "   --   $ " & STAT_INFORMATION(month_ind).stat_jobs_two_snap_pic_ave_inc_per_pay(each_memb) & " average inc/pay date"
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_three_job_counted(each_memb) = True Then
			Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_jobs_three_snap_pic_prosp_monthly_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_jobs_three_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 25, y_pos+10, 200, 10, "Paid " & STAT_INFORMATION(month_ind).stat_jobs_three_snap_pic_pay_freq(each_memb) & " - $ " & STAT_INFORMATION(month_ind).stat_jobs_three_snap_pic_ave_inc_per_pay(each_memb) & " average inc/pay date"
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_four_job_counted(each_memb) = True Then
			Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_jobs_four_snap_pic_prosp_monthly_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_jobs_four_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 25, y_pos+10, 200, 10, "Paid " & STAT_INFORMATION(month_ind).stat_jobs_four_snap_pic_pay_freq(each_memb) & " - $ " & STAT_INFORMATION(month_ind).stat_jobs_four_snap_pic_ave_inc_per_pay(each_memb) & " average inc/pay date"
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_five_job_counted(each_memb) = True Then
			Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_jobs_five_snap_pic_prosp_monthly_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_jobs_five_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 25, y_pos+10, 200, 10, "Paid " & STAT_INFORMATION(month_ind).stat_jobs_five_snap_pic_pay_freq(each_memb) & " - $ " & STAT_INFORMATION(month_ind).stat_jobs_five_snap_pic_ave_inc_per_pay(each_memb) & " average inc/pay date"
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_one_counted(each_memb) = True Then
			' Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - SELF EMP: " & STAT_INFORMATION(month_ind).stat_busi_one_type_info(each_memb) &
			Text 15, y_pos, 215, 10, "$ " & STAT_INFORMATION(month_ind).stat_busi_one_snap_prosp_net_inc(each_memb)& " - Monthly Income   --   Memb " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - SELF EMP: " & STAT_INFORMATION(month_ind).stat_busi_one_type_info(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_one_snap_income_verif_code(each_memb) = "N" or STAT_INFORMATION(month_ind).stat_busi_one_snap_expense_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 40, y_pos+10, 200, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_one_snap_prosp_gross_inc(each_memb) & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_one_snap_prosp_expenses(each_memb)
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_two_counted(each_memb) = True Then
			Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - SELF EMP: " & left(STAT_INFORMATION(month_ind).stat_busi_two_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_busi_two_snap_prosp_net_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_two_snap_income_verif_code(each_memb) = "N" or STAT_INFORMATION(month_ind).stat_busi_two_snap_expense_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 25, y_pos+10, 200, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_two_snap_prosp_gross_inc(each_memb) & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_two_snap_prosp_expenses(each_memb)
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_three_counted(each_memb) = True Then
			Text 15, y_pos, 215, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - SELF EMP: " & left(STAT_INFORMATION(month_ind).stat_busi_three_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_busi_three_snap_prosp_net_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_three_snap_income_verif_code(each_memb) = "N" or STAT_INFORMATION(month_ind).stat_busi_three_snap_expense_verif_code(each_memb) = "N" Then
				Text 40, y_pos+10, 200, 10, "Verification NOT Received."
			Else
				Text 25, y_pos+10, 200, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_three_snap_prosp_gross_inc(each_memb) & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_three_snap_prosp_expenses(each_memb)
			End If
			y_pos = y_pos + 20
		End If
		If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_one_counted(each_memb) = True Then
			Text 305, y_pos_2, 235, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_unea_one_snap_pic_prosp_monthly_inc(each_memb)
			y_pos_2 = y_pos_2 + 10
			If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then
				Text 330, y_pos_2, 200, 10, "Verification NOT Received."
				y_pos_2 = y_pos_2 + 10
			End If
		End If
		If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_two_counted(each_memb) = True Then
			Text 305, y_pos_2, 235, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_unea_two_snap_pic_prosp_monthly_inc(each_memb)
			y_pos_2 = y_pos_2 + 10
			If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then
				Text 330, y_pos_2, 200, 10, "Verification NOT Received."
				y_pos_2 = y_pos_2 + 10
			End If
		End If
		If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_three_counted(each_memb) = True Then
			Text 305, y_pos_2, 235, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_unea_three_snap_pic_prosp_monthly_inc(each_memb)
			y_pos_2 = y_pos_2 + 10
			If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then
				Text 330, y_pos_2, 200, 10, "Verification NOT Received."
				y_pos_2 = y_pos_2 + 10
			End If
		End If
		If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_four_counted(each_memb) = True Then
			Text 305, y_pos_2, 235, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_unea_four_snap_pic_prosp_monthly_inc(each_memb)
			y_pos_2 = y_pos_2 + 10
			If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then
				Text 330, y_pos_2, 200, 10, "Verification NOT Received."
				y_pos_2 = y_pos_2 + 10
			End If
		End If
		If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_five_counted(each_memb) = True Then
			Text 305, y_pos_2, 235, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & left(STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb) & "                              ", 30) & " Monhtly Income:   $ " & STAT_INFORMATION(month_ind).stat_unea_five_snap_pic_prosp_monthly_inc(each_memb)
			y_pos_2 = y_pos_2 + 10
			If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then
				Text 330, y_pos_2, 200, 10, "Verification NOT Received."
				y_pos_2 = y_pos_2 + 10
			End If
		End If
	  Next
	EndDialog
end function

function snap_elig_case_note()

	Call start_a_blank_case_note

	end_msg_info = end_msg_info & "NOTE entered for SNAP - " & elig_info & " eff " & first_month & header_end & vbCr
	Call write_variable_in_CASE_NOTE("APP Completed " & program_detail & " " & elig_info & " eff " & first_month & header_end)		'TODO - add closure or denial details here based on some other logic that we have to figure out'
	' If SNAP_ELIG_APPROVALS(approval).snap_elig_result = "ELIGIBLE" Then Call write_variable_in_CASE_NOTE("APP Completed - SNAP " & SNAP_ELIG_APPROVALS(approval).snap_elig_result & " eff " & first_month & " - Entitlement: $ " & SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot)
	' If SNAP_ELIG_APPROVALS(approval).snap_elig_result = "INELIGIBLE" Then Call write_variable_in_CASE_NOTE("APP Completed - SNAP " & SNAP_ELIG_APPROVALS(approval).snap_elig_result & " eff " & first_month)

	Call write_bullet_and_variable_in_CASE_NOTE("Approval completed", SNAP_ELIG_APPROVALS(elig_ind).snap_approved_date)
	' Call write_variable_in_CASE_NOTE("*** BENEFIT AMOUNT ***")
	If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "ELIGIBLE" Then
		Call write_variable_in_CASE_NOTE("================================ BENEFIT AMOUNT =============================")
		For approval = 0 to UBound(SNAP_ELIG_APPROVALS)
			If InStr(SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app), SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year) <> 0 Then
				If SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot = SNAP_ELIG_APPROVALS(approval).snap_benefit_amt Then
					' " 10/21:     Entitlement: $ 1,125.00 ¦ Issued to Resident: $ 1,125.00    10/21"
					Call write_variable_in_CASE_NOTE(" " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year & ": Entitlement: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot, 8) & "| Issued to Resident: $ " & right("        " & SNAP_ELIG_APPROVALS(approval).snap_benefit_amt, 8) & "         " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year)
				Else
					If SNAP_ELIG_APPROVALS(approval).snap_benefit_prorated_amt <> "" Then
						Call write_variable_in_CASE_NOTE(" " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year & ": Entitlement: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot, 8) & "|           Prorated: $ " & right("        " & SNAP_ELIG_APPROVALS(approval).snap_benefit_prorated_amt, 8) & "-from " & SNAP_ELIG_APPROVALS(approval).snap_benefit_prorated_date)
						If SNAP_ELIG_APPROVALS(approval).snap_benefit_amt_already_issued <> "" Then Call write_variable_in_CASE_NOTE("                               | Amt Already Issued: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_amt_already_issued, 8) & "  (-)")
						If SNAP_ELIG_APPROVALS(approval).snap_benefit_recoup_amount <> "0.00" Then Call write_variable_in_CASE_NOTE("                               |         Recoupment: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_recoup_amount, 8) & "  (-)")
					ElseIf SNAP_ELIG_APPROVALS(approval).snap_benefit_amt_already_issued <> "" Then
						Call write_variable_in_CASE_NOTE(" " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year & ": Entitlement: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot, 8) & "| Amt Already Issued: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_amt_already_issued, 8) & "  (-)")
						If SNAP_ELIG_APPROVALS(approval).snap_benefit_recoup_amount <> "0.00" Then Call write_variable_in_CASE_NOTE("                               |         Recoupment: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_recoup_amount, 8) & "  (-)")
					ElseIf SNAP_ELIG_APPROVALS(approval).snap_benefit_recoup_amount <> "0.00" Then
						Call write_variable_in_CASE_NOTE(" " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year & ": Entitlement: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot, 8) & "|         Recoupment: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_recoup_amount, 8) & "  (-)")
					End If

					Call write_variable_in_CASE_NOTE("                               | Issued to Resident: $ " & right("       " & SNAP_ELIG_APPROVALS(approval).snap_benefit_amt, 8) & "         " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year)
					Call write_variable_in_CASE_NOTE("                               |---------------------------------------------")
				End If
			End If
		Next
	End If

	' "======================================XX======================================"

	If SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = True Then
		Call write_variable_in_CASE_NOTE("============================= BUDGET FOR APPROVAL ===========================")

		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_prosp_gross_inc_test = "FAILED" Then Call write_variable_in_CASE_NOTE("SNAP INELIGIBLE because Prosp Inc exceeds GROSS INCOME MAX of $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_gross_inc)
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_prosp_net_inc_test = "FAILED" Then Call write_variable_in_CASE_NOTE("SNAP INELIGIBLE because Prosp Inc exceeds NET INCOME MAX of $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_net_adj_inc)
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_retro_gross_inc_test = "FAILED" Then Call write_variable_in_CASE_NOTE("SNAP INELIGIBLE because Retro Inc exceeds GROSS INCOME MAX of $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_gross_inc)
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_retro_net_inc_test = "FAILED" Then Call write_variable_in_CASE_NOTE("SNAP INELIGIBLE because Retro Inc exceeds NET INCOME MAX of $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_net_adj_inc)


		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc = "" Then SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc = "0.00"
		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_unea_inc = "" Then SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_unea_inc = "0.00"

		Call write_variable_in_CASE_NOTE(" SNAP Unit Size: " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_numb_in_assist_unit & " - Adult: " & SNAP_ELIG_APPROVALS(elig_ind).adults_recv_snap & ", Children: " & SNAP_ELIG_APPROVALS(elig_ind).children_recv_snap)

		beginning_txt = " Income:    "
		earned_info = "|   Gross Earned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc, 8)
		spaces_30 = "                              "
		For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
			If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_one_job_counted(each_memb) = True Then
				job_detail = left(STAT_INFORMATION(month_ind).stat_jobs_one_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb) & spaces_30, 26)
				Call write_variable_in_CASE_NOTE(beginning_txt & "Job- $" & job_detail & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_two_job_counted(each_memb) = True Then
				job_detail = left(STAT_INFORMATION(month_ind).stat_jobs_two_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb) & spaces_30, 26)
				Call write_variable_in_CASE_NOTE(beginning_txt & "Job- $" & job_detail & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_three_job_counted(each_memb) = True Then
				job_detail = left(STAT_INFORMATION(month_ind).stat_jobs_three_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb) & spaces_30, 26)
				Call write_variable_in_CASE_NOTE(beginning_txt & "Job- $" & job_detail & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_four_job_counted(each_memb) = True Then
				job_detail = left(STAT_INFORMATION(month_ind).stat_jobs_four_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb) & spaces_30, 26)
				Call write_variable_in_CASE_NOTE(beginning_txt & "Job- $" & job_detail & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_five_job_counted(each_memb) = True Then
				job_detail = left(STAT_INFORMATION(month_ind).stat_jobs_five_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb) & spaces_30, 26)
				Call write_variable_in_CASE_NOTE(beginning_txt & "Job- $" & job_detail & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_one_counted(each_memb) = True Then
				busi_details = left(STAT_INFORMATION(month_ind).stat_busi_one_snap_prosp_net_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_busi_one_type_info(each_memb) & spaces_30, 25)
				Call write_variable_in_CASE_NOTE(beginning_txt & "SELF- $" & busi_details & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_two_counted(each_memb) = True Then
				busi_details = left(STAT_INFORMATION(month_ind).stat_busi_two_snap_prosp_net_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_busi_two_type_info(each_memb) & spaces_30, 25)
				Call write_variable_in_CASE_NOTE(beginning_txt & "SELF- $" & busi_details & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_three_counted(each_memb) = True Then
				busi_details = left(STAT_INFORMATION(month_ind).stat_busi_three_snap_prosp_net_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb)  & "- " & STAT_INFORMATION(month_ind).stat_busi_three_type_info(each_memb) & spaces_30, 25)
				Call write_variable_in_CASE_NOTE(beginning_txt & "SELF- $" & busi_details & earned_info)
				beginning_txt = "            "
				earned_info = "|"
			End If
		Next
		If earned_info = "|   Gross Earned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc, 8) Then
			' Call write_variable_in_CASE_NOTE(" Income:    NONE                            |   Gross Earned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc, 8))
			Call write_variable_in_CASE_NOTE(" Income:    NO Earned Income                |   Gross Earned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_earned_inc, 8))

		End If
		unearned_info = "| Gross Unearned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_unea_inc, 8)
		For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
			If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_one_counted(each_memb) = True Then
				unea_detail = left(STAT_INFORMATION(month_ind).stat_unea_one_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & "- "& STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb), 31)
				Call write_variable_in_CASE_NOTE("            "  & "$" & unea_detail & unearned_info)
				unearned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_two_counted(each_memb) = True Then
				unea_detail = left(STAT_INFORMATION(month_ind).stat_unea_two_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & "- "& STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb), 31)
				Call write_variable_in_CASE_NOTE("            "  & "$" & unea_detail & unearned_info)
				unearned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_three_counted(each_memb) = True Then
				unea_detail = left(STAT_INFORMATION(month_ind).stat_unea_three_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & "- "& STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb), 31)
				Call write_variable_in_CASE_NOTE("            "  & "$" & unea_detail & unearned_info)
				unearned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_four_counted(each_memb) = True Then
				unea_detail = left(STAT_INFORMATION(month_ind).stat_unea_four_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & "- "& STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb), 31)
				Call write_variable_in_CASE_NOTE("            "  & "$" & unea_detail & unearned_info)
				unearned_info = "|"
			End If
			If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_five_counted(each_memb) = True Then
				unea_detail = left(STAT_INFORMATION(month_ind).stat_unea_five_snap_pic_prosp_monthly_inc(each_memb) & "- M" & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & "- "& STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb), 31)
				Call write_variable_in_CASE_NOTE("            "  & "$" & unea_detail & unearned_info)
				unearned_info = "|"
			End If
		Next
		If unearned_info = "| Gross Unearned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_unea_inc, 8) Then Call write_variable_in_CASE_NOTE("            NO Unearned Income              | Gross Unearned Inc: $" & right("        "&SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_unea_inc, 8))

		Call write_variable_in_CASE_NOTE("                                            |    Total Gross Inc: $" & right("        " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_gross_inc, 8))

		deduction_detail_one = ""
		deduction_detail_two = ""
		deduction_detail_three = ""

		' Standard $177
		' Earned Inc $126
		' Medical Exp $0
		' Depndt Care $175.00
		' 1234567890123456789
		' Child Suprt $0

		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard <> "" Then
			If deduction_detail_one = "" Then
				deduction_detail_one = left("Standard $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_one) < 21 Then
				deduction_detail_one = deduction_detail_one & "- Standard $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard, ".00", "")
			ElseIf deduction_detail_two = "" Then
				deduction_detail_two = left("Standard $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_two) < 21 Then
				deduction_detail_two = deduction_detail_two & "- Standard  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard, ".00", "")
			ElseIf deduction_detail_three = "" Then
				deduction_detail_three = deduction_detail_three & "Standard $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard, ".00", "")
			ElseIf len(deduction_detail_three) < 21 Then
				deduction_detail_three = deduction_detail_three & "- Standard  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_standard, ".00", "")
			End if
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned <> "" Then
			' MsgBox SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned
			If deduction_detail_one = "" Then
				deduction_detail_one = left("Earned Inc $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_one) < 21 Then
				deduction_detail_one = deduction_detail_one & "- Earned Inc $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned, ".00", "")
			ElseIf deduction_detail_two = "" Then
				deduction_detail_two = left("Earned Inc $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_two) < 21 Then
				deduction_detail_two = deduction_detail_two & "- Earned Inc $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned, ".00", "")
			ElseIf deduction_detail_three = "" Then
				deduction_detail_three = deduction_detail_three & "Earned Inc $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned, ".00", "")
			ElseIf len(deduction_detail_three) < 21 Then
				deduction_detail_three = deduction_detail_three & "- Earned Inc  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_earned, ".00", "")
			End if
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical <> "" Then
			If deduction_detail_one = "" Then
				deduction_detail_one = left("Medical Exp $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_one) < 21 Then
				deduction_detail_one = deduction_detail_one & "- Medical Exp $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical, ".00", "")
			ElseIf deduction_detail_two = "" Then
				deduction_detail_two = left("Medical Exp $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_two) < 21 Then
				deduction_detail_two = deduction_detail_two & "- Medical Exp $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical, ".00", "")
			ElseIf deduction_detail_three = "" Then
				deduction_detail_three = deduction_detail_three & "Medical Exp $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical, ".00", "")
			ElseIf len(deduction_detail_three) < 21 Then
				deduction_detail_three = deduction_detail_three & "- Medical Exp  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_medical, ".00", "")
			End if
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care <> "" Then
			If deduction_detail_one = "" Then
				deduction_detail_one = left("Depndt Care $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_one) < 21 Then
				deduction_detail_one = deduction_detail_one & "-Depndt Care  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care, ".00", "")
			ElseIf deduction_detail_two = "" Then
				deduction_detail_two = left("Depndt Care $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_two) < 21 Then
				deduction_detail_two = deduction_detail_two & "- Depndt Care $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care, ".00", "")
			ElseIf deduction_detail_three = "" Then
				deduction_detail_three = deduction_detail_three & "Depndt Care $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care, ".00", "")
			ElseIf len(deduction_detail_three) < 21 Then
				deduction_detail_three = deduction_detail_three & "- Depndt Care  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_depndt_care, ".00", "")
			End if
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses <> "" Then
			If deduction_detail_one = "" Then
				deduction_detail_one = left("Child Suprt $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_one) < 21 Then
				deduction_detail_one = deduction_detail_one & "- Child Suprt $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses, ".00", "")
			ElseIf deduction_detail_two = "" Then
				deduction_detail_two = left("Child Suprt $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses, ".00", "") & spaces_18, 15)
			ElseIf len(deduction_detail_two) < 21 Then
				deduction_detail_two = deduction_detail_two & "- Child Suprt $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses, ".00", "")
			ElseIf deduction_detail_three = "" Then
				deduction_detail_three = deduction_detail_three & "Child Suprt $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses, ".00", "")
			ElseIf len(deduction_detail_three) < 21 Then
				deduction_detail_three = deduction_detail_three & "- Child Suprt  $" & replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_deduct_cses, ".00", "")
			End if
		End If
		' MsgBox "deduction_detail_one - " & deduction_detail_one & vbCr &_
		' 		"deduction_detail_two - " & deduction_detail_two & vbCr &_
		' 		"deduction_detail_three - " & deduction_detail_three
		Call write_variable_in_CASE_NOTE(" Deductions:" & left(deduction_detail_one & spaces_30, 32) & "|   (-)   Deductions: $" & right("        " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_deduct, 8))
		If deduction_detail_two <> "" Then Call write_variable_in_CASE_NOTE("            " & left(deduction_detail_two & spaces_30, 32) & "|")
		If deduction_detail_three <> "" Then Call write_variable_in_CASE_NOTE("            " & left(deduction_detail_three & spaces_30, 32) & "|")
		' MsgBox "CHECK"
		Call write_variable_in_CASE_NOTE("                                            |            Net Inc: $" & right("        " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_net_inc, 8))
		Call write_variable_in_CASE_NOTE(" Expenses:  Housing: $"& left(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_housing_exp_total & "        ", 8)& "        |--------------------------|")
		' " Expenses:  Housing: $1,250.00        |-------------------------------------|"


		Call write_variable_in_CASE_NOTE("            Utilities: $"& left(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_utilities_exp_total&"        ", 8) & "      | Total Shelter: $ " & left(replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_shel_total, ".00", "")&"        ", 8) & "|")
		Call write_variable_in_CASE_NOTE("            MAX Allowable: $"& left(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_max_allow_shel&"        ", 8) & "  |(-)1/2 Net Inc: $ " & left(replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_50_perc_net_inc, ".00", "")&"        ", 8) & "|")
		Call write_variable_in_CASE_NOTE("                                      |   Adj Shelter: $ " & left(replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_adj_shel_costs, ".00", "")&"        ", 8) & "|")
		Call write_variable_in_CASE_NOTE("                                      |--------------------------|")
		Call write_variable_in_CASE_NOTE("                                            |  (-)Allow Shel Exp: $" & right("        " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_shel_expenses, 8))
		Call write_variable_in_CASE_NOTE("                                            |Net Adjusted Income: $" & right("        " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_net_adj_inc, 8))
		Call write_variable_in_CASE_NOTE("                               |---------------------------------|")
		Call write_variable_in_CASE_NOTE("                               |    Thrifty Food Plan: $ " & left(replace(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_thrifty_food_plan, ".00", "")&"        ", 8) & "|")
		Call write_variable_in_CASE_NOTE("                               |(-)30% of Net Adj Inc: $ " & left(replace(SNAP_ELIG_APPROVALS(elig_ind).snap_bug_30_percent_net_adj_inc, ".00", "")&"        ", 8) & "|")
		Call write_variable_in_CASE_NOTE("                               |---------------------------------|")
		Call write_variable_in_CASE_NOTE("                                            |   SNAP Entitlement: $" & right("        " & SNAP_ELIG_APPROVALS(elig_ind).snap_benefit_monthly_fs_allot, 8))

		If SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_sent, unique_app) = True Then
			Call write_variable_in_CASE_NOTE("SNAP Budgeted Gross Income of  $ " & SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_gross_inc & " exceeds 130% FPG of $ " & FormatNumber(SNAP_UNIQUE_APPROVALS(snap_130_percent_fpg_amt, unique_app), 2, -1, 0, -1))

			If SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_sent, unique_app) = True Then Call write_variable_in_CASE_NOTE(" - WCOM added to Notice for to clarify reporting responsibilities.")
		End If
		' Call write_variable_in_CASE_NOTE("*** CASE STATUS ***")
	End If

	If SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False Then
		Call write_variable_in_CASE_NOTE("================================== CASE TESTS ===============================")

		Call write_variable_in_CASE_NOTE("* SNAP is INELIGIBLE because not all CASE TESTS were passed.") '' to make this Household Eligible")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_appl_withdrawn_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - The request for SNAP benefits was withdrawn. (APPLICATION WITHDRAWN)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_applct_elig_test = "FAILED" Then
			Call write_variable_in_CASE_NOTE(" - The applicant is not SNAP eligibile. (APPLICANT ELIGIBLE)")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_abawd(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has reached the SNAP Time Limit - ABAWD")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_absence(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is not in the household.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_roomer(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is a roomer.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_boarder(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is a boarder.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_citizenship(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 does not meet citizenship requirements.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_citizenship_coop(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has not complied with citizzenship information.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_cmdty(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has received commodities for this time period.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_disq(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is disqualified from SNAP")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_dupl_assist(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has received SNAP assisnce on another case.")

			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_fraud(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has a Fraud determination.")
			If STAT_INFORMATION(month_ind).stat_disq_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_one_program(each_memb) = "SNAP" Then
				Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_one_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_one_begin_date(each_memb))
				If IsDate(STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb))
			End If
			If STAT_INFORMATION(month_ind).stat_disq_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_two_program(each_memb) = "SNAP" Then
				Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_two_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_two_begin_date(each_memb))
				If IsDate(STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb))
			End If
			If STAT_INFORMATION(month_ind).stat_disq_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_three_program(each_memb) = "SNAP" Then
				Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_three_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_three_begin_date(each_memb))
				If IsDate(STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb))
			End If
			If STAT_INFORMATION(month_ind).stat_disq_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_four_program(each_memb) = "SNAP" Then
				Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_four_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_four_begin_date(each_memb))
				If IsDate(STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb))
			End If
			If STAT_INFORMATION(month_ind).stat_disq_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_five_program(each_memb) = "SNAP" Then
				Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_five_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_five_begin_date(each_memb))
				If IsDate(STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb))
			End If

			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_eligible_student(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is an ineligible student.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_institution(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is in an institution.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_mfip_elig(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is MFIP eligible.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_non_applcnt(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is not requesting SNAP.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_residence(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 does not have MN residence.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_ssn_coop(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has not cooperated with SSN requirements.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_unit_memb(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 is not a unit member.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_work_reg(0) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb 01 has not complied with work registration.")

		End If

		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_comdty_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - The case has received Commodity Food in this time period. (COMMODITY)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_disq_test = "FAILED" Then
			Call write_variable_in_CASE_NOTE(" - This case has a Disqualification. (DISQUALIFICATION)")
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_disq_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_one_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_one_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_one_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_two_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_two_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_two_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_three_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_three_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_three_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_four_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_four_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_four_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_five_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_five_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_five_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb))
				End If
			Next
		End If

		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_dupl_assist_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - This case has already received SNAP. (DUPL ASSISTANCE)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_eligible_person_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - There is not eligible person on this case. (ELIGIBLE PERSON)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_fail_coop_test = "FAILED" Then
			Call write_variable_in_CASE_NOTE(" - This case has failed to cooperate. (FAIL TO COOPERATE)")
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_disq_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_source(each_memb) = "NON-COOP" AND STAT_INFORMATION(month_ind).stat_disq_one_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_one_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_one_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_source(each_memb) = "NON-COOP" AND STAT_INFORMATION(month_ind).stat_disq_two_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_two_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_two_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_source(each_memb) = "NON-COOP" AND STAT_INFORMATION(month_ind).stat_disq_three_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_three_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_three_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_source(each_memb) = "NON-COOP" AND STAT_INFORMATION(month_ind).stat_disq_four_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_four_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_four_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_source(each_memb) = "NON-COOP" AND STAT_INFORMATION(month_ind).stat_disq_five_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_five_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_five_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb))
				End If
			Next
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_fail_file_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - This case has failed to file a report. (FAIL TO FILE)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_fail_file_hrf = "FAILED" Then Call write_variable_in_CASE_NOTE("    -Monthly Household Report process was not completed.")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_fail_file_sr = "FAILED" Then Call write_variable_in_CASE_NOTE("    -Six Month Report process was not completed.")

		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_recert_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - The annual recertification process was not completed. (RECERTIFICATION)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_residence_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - This case has not established Minnesota residency. (RESIDENCE)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_resource_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - The assets have exceeded the max. (RESOURCE)")
		' TODO - add more asset test information'
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_strike_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - The case has a member on strike. (STRIKE)")
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_xfer_resource_inc_test = "FAILED" Then
			Call write_variable_in_CASE_NOTE(" - This case has failed transfer resources/income. (TRANSFER RESOURCE INC)")
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_disq_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_source(each_memb) = "TRANSFER" AND STAT_INFORMATION(month_ind).stat_disq_one_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_one_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_one_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_source(each_memb) = "TRANSFER" AND STAT_INFORMATION(month_ind).stat_disq_two_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_two_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_two_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_source(each_memb) = "TRANSFER" AND STAT_INFORMATION(month_ind).stat_disq_three_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_three_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_three_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_source(each_memb) = "TRANSFER" AND STAT_INFORMATION(month_ind).stat_disq_four_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_four_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_four_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_source(each_memb) = "TRANSFER" AND STAT_INFORMATION(month_ind).stat_disq_five_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_five_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_five_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb))
				End If
			Next
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test = "FAILED" Then
			Call write_variable_in_CASE_NOTE(" - Verifications were not received. (VERIFICATION)")
			Call write_variable_in_CASE_NOTE("   VERIFICATION REQUEST FORM SENT: " & SNAP_UNIQUE_APPROVALS(verif_reqquest_date, unique_app) & ", due by: " & due_date)
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_MEMB_ID  = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Proof of the identity of the Applicant was not received.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_ACCT = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Proof of bank account not received.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_PACT = "FAILED" Then
				Call write_variable_in_CASE_NOTE("   - Case ineligible due to: " & SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, unique_app) & ". ")
				Call write_variable_in_CASE_NOTE("     INELIG created using PACT.")
				If SNAP_UNIQUE_APPROVALS(pact_wcom_sent, unique_app) = True Then Call write_variable_in_CASE_NOTE("     WCOM added to Notice with PACT reason.")
			End If
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_ADDR = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Residency not verified.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_SECU = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Proof of securities not received.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_RBIC = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Self Employment - Roomer/Boarder Income not verified.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_BUSI = "FAILED" Then
				Call write_variable_in_CASE_NOTE("   - Self Employment income not verified.")
				For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
					If STAT_INFORMATION(month_ind).stat_busi_one_snap_income_verif_code(each_memb) = "N" or STAT_INFORMATION(month_ind).stat_busi_one_snap_expense_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " Self Employment verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_busi_two_snap_income_verif_code(each_memb) = "N" or STAT_INFORMATION(month_ind).stat_busi_two_snap_expense_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " Self Employment verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_busi_three_snap_income_verif_code(each_memb) = "N" or STAT_INFORMATION(month_ind).stat_busi_three_snap_expense_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " Self Employment verif not received.")
					End if
				Next
			End If
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_SPON = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Sponsor income not verified.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_STIN = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Student income not verified.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_UNEA = "FAILED" Then
				Call write_variable_in_CASE_NOTE("   - Unearned income not verified.")
				For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
					If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb) & " verif not received.")
					End if
				Next
			End If
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_JOBS = "FAILED" Then
				Call write_variable_in_CASE_NOTE("   - Wage income not verified.")
				For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
					If STAT_INFORMATION(month_ind).stat_jobs_one_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " employment at " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_jobs_two_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " employment at " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_jobs_three_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " employment at " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_jobs_four_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " employment at " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb) & " verif not received.")
					End if
					If STAT_INFORMATION(month_ind).stat_jobs_five_verif_code(each_memb) = "N" Then
						Call write_variable_in_CASE_NOTE("     M " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " employment at " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb) & " verif not received.")
					End if
				Next
			End If
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_STWK = "FAILED" Then Call write_variable_in_CASE_NOTE("   - End of employment not verified.")
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_STRK = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Proof of strike was not received.")
		End If

		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_voltry_quit_test = "FAILED" Then
			Call write_variable_in_CASE_NOTE(" - This case has a member who quit work, not following SNAP general work rules. (VOLUNTARY QUIT)")
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_disq_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_source(each_memb) = "VOL QUIT" AND STAT_INFORMATION(month_ind).stat_disq_one_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_one_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_one_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_source(each_memb) = "VOL QUIT" AND STAT_INFORMATION(month_ind).stat_disq_two_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_two_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_two_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_source(each_memb) = "VOL QUIT" AND STAT_INFORMATION(month_ind).stat_disq_three_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_three_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_three_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_source(each_memb) = "VOL QUIT" AND STAT_INFORMATION(month_ind).stat_disq_four_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_four_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_four_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_source(each_memb) = "VOL QUIT" AND STAT_INFORMATION(month_ind).stat_disq_five_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_five_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_five_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb))
				End If
			Next
		End If
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_work_reg_test = "FAILED" Then Call write_variable_in_CASE_NOTE(" - The persons on this case did not comply with work registration. (WORK REGISTRATION)")
		' TODO - add more detail for work Reg'
	End If

	'312524'
	first_memb = ""
	If SNAP_ELIG_APPROVALS(elig_ind).snap_case_applct_elig_test = "FAILED" and UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb) > 0 Then first_memb = 1
	If SNAP_ELIG_APPROVALS(elig_ind).snap_case_applct_elig_test <> "FAILED" Then first_memb = 0

	first_inelig_memb = True
	If first_memb <> "" Then
		For each_memb = first_memb to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)

			If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_eligibility(each_memb) = "INELIGIBLE" Then

				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "ELIGIBLE" Then
					If first_inelig_memb = True Then
						Call write_variable_in_CASE_NOTE("================================ MEMBER TESTS ===============================")
						first_inelig_memb = False
					End If
					Call write_variable_in_CASE_NOTE(" - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is not eligible for SNAP and not included in the SNAP benefit.")
					Call write_variable_in_CASE_NOTE("   The income for this member is " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_counted(each_memb))
				ElseIf SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_failed_test(each_memb) = True Then
					If first_inelig_memb = True Then
						Call write_variable_in_CASE_NOTE("================================ MEMBER TESTS ===============================")
						first_inelig_memb = False
					End If
					Call write_variable_in_CASE_NOTE(" - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is not eligible for SNAP.")
				End If

				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_abawd(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has reached the SNAP Time Limit - ABAWD")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_absence(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is not in the household.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_roomer(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is a roomer.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_boarder(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is a boarder.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_citizenship(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " does not meet citizenship requirements.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_citizenship_coop(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has not complied with citizzenship information.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_cmdty(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has received commodities for this time period.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_disq(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is disqualified from SNAP")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_dupl_assist(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has received SNAP assisnce on another case.")

				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_fraud(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has a Fraud determination.")
				If STAT_INFORMATION(month_ind).stat_disq_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_one_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_one_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_one_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_one_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_one_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_two_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_two_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_two_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_two_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_two_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_three_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_three_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_three_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_three_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_three_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_four_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_four_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_four_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_four_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_four_end_date(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_disq_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_active(each_memb) = True AND STAT_INFORMATION(month_ind).stat_disq_five_source(each_memb) = "DISQUAL" AND STAT_INFORMATION(month_ind).stat_disq_five_program(each_memb) = "SNAP" Then
					Call write_variable_in_CASE_NOTE("   - " & STAT_INFORMATION(month_ind).stat_disq_five_type_info(each_memb) & " begin date: " & STAT_INFORMATION(month_ind).stat_disq_five_begin_date(each_memb))
					If IsDate(STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb)) = True Then Call write_variable_in_CASE_NOTE("     Disqualification to end on " & STAT_INFORMATION(month_ind).stat_disq_five_end_date(each_memb))
				End If

				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_eligible_student(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is an ineligible student.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_institution(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is in an institution.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_mfip_elig(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is MFIP eligible.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_non_applcnt(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is not requesting SNAP.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_residence(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " does not have MN residence.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_ssn_coop(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has not cooperated with SSN requirements.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_unit_memb(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " is not a unit member.")
				If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_membs_work_reg(each_memb) = "FAILED" Then Call write_variable_in_CASE_NOTE("   - Memb " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_ref_numbs(each_memb) & " has not complied with work registration.")
			End If
		Next
	End if

	Call write_variable_in_CASE_NOTE("================================= CASE STATUS ===============================")
	spaces_18 = "                  "
	' Call write_variable_in_CASE_NOTE(" SNAP Status:      " & left(SNAP_ELIG_APPROVALS(elig_ind).snap_curr_prog_status & spaces_18, 18) & "Budget Cycle:     " & SNAP_ELIG_APPROVALS(elig_ind).snap_budget_cycle)
	' Call write_variable_in_CASE_NOTE(" Reporting Status: " & left(SNAP_ELIG_APPROVALS(elig_ind).snap_reporting_status & spaces_18, 18) & "Review Date:      " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_revw_date)
	Call write_variable_in_CASE_NOTE("SNAP Status:      " & SNAP_ELIG_APPROVALS(elig_ind).snap_curr_prog_status)
	Call write_variable_in_CASE_NOTE("Budget Cycle:     " & SNAP_ELIG_APPROVALS(elig_ind).snap_budget_cycle)
	If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "ELIGIBLE" Then
		Call write_variable_in_CASE_NOTE("Reporting Status: " & SNAP_ELIG_APPROVALS(elig_ind).snap_reporting_status)
		Call write_variable_in_CASE_NOTE("Review Date:      " & SNAP_ELIG_APPROVALS(elig_ind).snap_elig_revw_date)
	End If
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	' MsgBox SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app)
	PF3
end function


'DECLARATIONS===============================================================================================================
class dwp_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found

	public dwp_elig_ref_numbs()
	public dwp_elig_membs_full_name()
	public dwp_elig_membs_request_yn()
	public dwp_elig_membs_member_code()
	public dwp_elig_membs_member_info()
	public dwp_elig_membs_funding_source_code()
	public dwp_elig_membs_funding_source_info()
	public dwp_elig_membs_elig_status()
	public dwp_elig_membs_begin_date()
	public dwp_elig_membs_adult_or_child()
	public dwp_elig_membs_test_absence()
	public dwp_elig_membs_test_child_age()
	public dwp_elig_membs_test_citizenship()
	public dwp_elig_membs_test_citizenship_verif()
	public dwp_elig_membs_test_dupl_assistance()
	public dwp_elig_membs_test_foster_care()
	public dwp_elig_membs_test_fraud()
	public dwp_elig_membs_test_minor_living_arrangement()
	public dwp_elig_membs_test_post_60_removal()
	public dwp_elig_membs_test_ssi()
	public dwp_elig_membs_test_ssn_coop()
	public dwp_elig_membs_test_unit_member()
	public dwp_elig_membs_test_unlawful_conduct()
	public dwp_elig_membs_es_status_code()
	public dwp_elig_membs_es_status_info()

	public dwp_elig_case_test_application_withdrawn
	public dwp_elig_case_test_assets
	public dwp_elig_case_test_CS_disqualification
	public dwp_elig_case_test_death_of_applicant
	public dwp_elig_case_test_dupl_assistance
	public dwp_elig_case_test_eligible_child
	public dwp_elig_case_test_ES_disqualification
	public dwp_elig_case_test_fail_coop
	public dwp_elig_case_test_four_month_limit
	public dwp_elig_case_test_initial_income
	public dwp_elig_case_test_MFIP_conversion
	public dwp_elig_case_test_residence
	public dwp_elig_case_test_strike
	public dwp_elig_case_test_TANF_time_limit
	public dwp_elig_case_test_transfer_of_assets
	public dwp_elig_case_test_verif
	public dwp_elig_case_test_new_spouse_income
	public dwp_elig_asset_CASH
	public dwp_elig_asset_ACCT
	public dwp_elig_asset_SECU
	public dwp_elig_asset_CARS
	public dwp_elig_asset_SPON
	public dwp_elig_asset_total
	public dwp_elig_asset_maximum
	public dwp_elig_test_fail_coop_applied_other_benefits
	public dwp_elig_test_fail_coop_provide_requested_info
	public dwp_elig_test_fail_coop_IEVS
	public dwp_elig_test_fail_coop_vendor_info
	public dwp_elig_initial_counted_earned_income
	public dwp_elig_initial_dependent_care_expense
	public dwp_elig_initial_counted_unearned_incom
	public dwp_elig_initial_counted_deemed_income
	public dwp_elig_initial_child_support_exclusion
	public dwp_elig_initial_total_counted_income
	public dwp_elig_initial_family_wage_level
	public dwp_elig_test_verif_ACCT
	public dwp_elig_test_verif_BUSI
	public dwp_elig_test_verif_CARS
	public dwp_elig_test_verif_JOBS
	public dwp_elig_test_verif_MEMB_dob
	public dwp_elig_test_verif_MEMB_id
	public dwp_elig_test_verif_PARE
	public dwp_elig_test_verif_PREG
	public dwp_elig_test_verif_RBIC
	public dwp_elig_test_verif_ADDR
	public dwp_elig_test_verif_SCHL
	public dwp_elig_test_verif_SECU
	public dwp_elig_test_verif_SPON
	public dwp_elig_test_verif_UNEA

	public dwp_elig_budg_shel_rent_mortgage
	public dwp_elig_budg_shel_property_tax
	public dwp_elig_budg_shel_house_insurance
	public dwp_elig_budg_hest_electricity
	public dwp_elig_budg_hest_heat_air
	public dwp_elig_budg_hest_water_sewer_garbage
	public dwp_elig_budg_hest_phone
	public dwp_elig_budg_shel_other
	public dwp_elig_budg_total_shelter_costs
	public dwp_elig_budg_personal_needs
	public dwp_elig_budg_total_DWP_need
	public dwp_elig_budg_earned_income
	public dwp_elig_budg_unearned_income
	public dwp_elig_budg_deemed_income
	public dwp_elig_budg_child_support_exclusion
	public dwp_elig_budg_budget_month_total
	public dwp_elig_budg_prior_low
	public dwp_elig_budg_DWP_countable_income
	public dwp_elig_budg_unmet_need
	public dwp_elig_budg_DWP_max_grant
	public dwp_elig_budg_DWP_grant
	public dwp_elig_cses_income
	public dwp_elig_child_count

	public dwp_elig_prorated_date
	public dwp_elig_prorated_amount
	public dwp_elig_amount_already_issued
	public dwp_elig_supplement_due
	public dwp_elig_overpayment
	public dwp_elig_adjusted_grant_amount
	public dwp_elig_recoupment_amount
	public dwp_elig_shelter_benefit_grant
	public dwp_elig_personal_needs_grant
	public dwp_elig_overpayment_fed_hh_count
	public dwp_elig_overpayment_fed_amount
	public dwp_elig_overpayment_state_hh_count
	public dwp_elig_overpayment_state_amount
	public dwp_elig_adjusted_grant_fed_hh_count
	public dwp_elig_adjusted_grant_fed_amount
	public dwp_elig_adjusted_grant_state_hh_count
	public dwp_elig_adjusted_grant_state_amount

	public dwp_approved_date
	public dwp_process_date
	public dwp_prev_approval
	public dwp_case_last_approval_date
	public dwp_case_current_prog_status
	public dwp_case_eligibility_result
	public dwp_case_source_of_info
	public dwp_case_benefit_impact
	public dwp_case_4th_month_of_elig
	public dwp_case_caregivers_have_es_plan
	public dwp_case_responsible_county
	public dwp_case_service_county
	public dwp_case_asst_unit_caregivers
	public dwp_case_asst_unit_children
	public dwp_case_total_assets
	public dwp_case_maximum_assets
	public dwp_case_summary_grant_amount
	public dwp_case_summary_net_grant_amount
	public dwp_case_summary_shelter_benefit_portion
	public dwp_case_summary_personal_needs_portion


	public sub read_elig()
		call navigate_to_MAXIS_screen("ELIG", "DWP ")
		EMWriteScreen elig_footer_month, 20, 56
		EMWriteScreen elig_footer_year, 20, 59
		approved_today = False
		approved_version_found = False
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

			ReDim dwp_elig_ref_numbs(0)
			ReDim dwp_elig_membs_full_name(0)
			ReDim dwp_elig_membs_request_yn(0)
			ReDim dwp_elig_membs_member_code(0)
			ReDim dwp_elig_membs_member_info(0)
			ReDim dwp_elig_membs_funding_source_code(0)
			ReDim dwp_elig_membs_funding_source_info(0)
			ReDim dwp_elig_membs_elig_status(0)
			ReDim dwp_elig_membs_begin_date(0)
			ReDim dwp_elig_membs_adult_or_child(0)
			ReDim dwp_elig_membs_test_absence(0)
			ReDim dwp_elig_membs_test_child_age(0)
			ReDim dwp_elig_membs_test_citizenship(0)
			ReDim dwp_elig_membs_test_citizenship_verif(0)
			ReDim dwp_elig_membs_test_dupl_assistance(0)
			ReDim dwp_elig_membs_test_foster_care(0)
			ReDim dwp_elig_membs_test_fraud(0)
			ReDim dwp_elig_membs_test_minor_living_arrangement(0)
			ReDim dwp_elig_membs_test_post_60_removal(0)
			ReDim dwp_elig_membs_test_ssi(0)
			ReDim dwp_elig_membs_test_ssn_coop(0)
			ReDim dwp_elig_membs_test_unit_member(0)
			ReDim dwp_elig_membs_test_unlawful_conduct(0)
			ReDim dwp_elig_membs_es_status_code(0)
			ReDim dwp_elig_membs_es_status_info(0)

			row = 7
			elig_memb_count = 0
			Do
				EMReadScreen ref_numb, 2, row, 5

				ReDim preserve dwp_elig_ref_numbs(elig_memb_count)
				ReDim preserve dwp_elig_membs_full_name(elig_memb_count)
				ReDim preserve dwp_elig_membs_request_yn(elig_memb_count)
				ReDim preserve dwp_elig_membs_member_code(elig_memb_count)
				ReDim preserve dwp_elig_membs_member_info(elig_memb_count)
				ReDim preserve dwp_elig_membs_funding_source_code(elig_memb_count)
				ReDim preserve dwp_elig_membs_funding_source_info(elig_memb_count)
				ReDim preserve dwp_elig_membs_elig_status(elig_memb_count)
				ReDim preserve dwp_elig_membs_begin_date(elig_memb_count)
				ReDim preserve dwp_elig_membs_adult_or_child(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_absence(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_child_age(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_citizenship(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_citizenship_verif(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_dupl_assistance(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_foster_care(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_fraud(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_minor_living_arrangement(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_post_60_removal(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_ssi(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_ssn_coop(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_unit_member(elig_memb_count)
				ReDim preserve dwp_elig_membs_test_unlawful_conduct(elig_memb_count)
				ReDim preserve dwp_elig_membs_es_status_code(elig_memb_count)
				ReDim preserve dwp_elig_membs_es_status_info(elig_memb_count)

				dwp_elig_ref_numbs(elig_memb_count) = ref_numb
				EMReadScreen full_name_information, 20, row, 9
				full_name_information = trim(full_name_information)
				name_array = split(full_name_information, " ")
				For each name_parts in name_array
					If len(name_parts) <> 1 Then dwp_elig_membs_full_name(elig_memb_count) = dwp_elig_membs_full_name(elig_memb_count) & " " & name_parts
				Next
				dwp_elig_membs_full_name(elig_memb_count) = trim((dwp_elig_membs_full_name(elig_memb_count)))

				EMReadScreen dwp_elig_membs_request_yn(elig_memb_count), 1, row, 31
				EMReadScreen dwp_elig_membs_member_code(elig_memb_count), 1, row, 35
				EMReadScreen dwp_elig_membs_funding_source_code(elig_memb_count), 1, row, 53
				EMReadScreen dwp_elig_membs_elig_status(elig_memb_count), 12, row, 57
				EMReadScreen dwp_elig_membs_begin_date(elig_memb_count), 8, row, 73

				dwp_elig_membs_elig_status(elig_memb_count) = trim(dwp_elig_membs_elig_status(elig_memb_count))

				If dwp_elig_membs_member_code(elig_memb_count) = "A" Then dwp_elig_membs_member_info(elig_memb_count) = "Eligible"
				If dwp_elig_membs_member_code(elig_memb_count) = "D" Then dwp_elig_membs_member_info(elig_memb_count) = "SSI/IVE/Adoption Assistance Recipient"
				If dwp_elig_membs_member_code(elig_memb_count) = "F" Then dwp_elig_membs_member_info(elig_memb_count) = "Ineligible, Deemer"
				If dwp_elig_membs_member_code(elig_memb_count) = "G" Then dwp_elig_membs_member_info(elig_memb_count) = "Parent of Minor Caregiver, Deemer"
				If dwp_elig_membs_member_code(elig_memb_count) = "H" Then dwp_elig_membs_member_info(elig_memb_count) = "Other Deemer"
				If dwp_elig_membs_member_code(elig_memb_count) = "I" Then dwp_elig_membs_member_info(elig_memb_count) = "Ineligible, Pare of Unit"
				If dwp_elig_membs_member_code(elig_memb_count) = "J" Then dwp_elig_membs_member_info(elig_memb_count) = "Ineligible, Deemer"
				If dwp_elig_membs_member_code(elig_memb_count) = "N" Then dwp_elig_membs_member_info(elig_memb_count) = "Not Counted"

				If dwp_elig_membs_funding_source_code(elig_memb_count) = "F" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "Federal Funds (TANF Cash)"
				If dwp_elig_membs_funding_source_code(elig_memb_count) = "S" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "State Funds (Cash)"
				If dwp_elig_membs_funding_source_code(elig_memb_count) = "I" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "Ineligible for DWP"
				If dwp_elig_membs_funding_source_code(elig_memb_count) = "N" Then dwp_elig_membs_funding_source_info(elig_memb_count) = "Not Applicable"

				Call write_value_and_transmit("X", row, 3)		'open member test information
				EMReadScreen dwp_elig_membs_adult_or_child(elig_memb_count), 1, 7, 51

				If dwp_elig_membs_adult_or_child(elig_memb_count) = "A" Then dwp_elig_membs_adult_or_child(elig_memb_count) = "Adult"
				If dwp_elig_membs_adult_or_child(elig_memb_count) = "C" Then dwp_elig_membs_adult_or_child(elig_memb_count) = "Child"

				EMReadScreen dwp_elig_membs_test_absence(elig_memb_count), 			6, 10, 7
				EMReadScreen dwp_elig_membs_test_child_age(elig_memb_count), 		6, 11, 7
				EMReadScreen dwp_elig_membs_test_citizenship(elig_memb_count), 		6, 12, 7
				EMReadScreen dwp_elig_membs_test_citizenship_verif(elig_memb_count), 6, 13, 7
				EMReadScreen dwp_elig_membs_test_dupl_assistance(elig_memb_count), 	6, 14, 7
				EMReadScreen dwp_elig_membs_test_foster_care(elig_memb_count), 		6, 15, 7
				EMReadScreen dwp_elig_membs_test_fraud(elig_memb_count), 			6, 16, 7

				EMReadScreen dwp_elig_membs_test_minor_living_arrangement(elig_memb_count), 6, 10, 43
				EMReadScreen dwp_elig_membs_test_post_60_removal(elig_memb_count), 			6, 11, 43
				EMReadScreen dwp_elig_membs_test_ssi(elig_memb_count), 						6, 12, 43
				EMReadScreen dwp_elig_membs_test_ssn_coop(elig_memb_count), 				6, 13, 43
				EMReadScreen dwp_elig_membs_test_unit_member(elig_memb_count), 				6, 14, 43
				EMReadScreen dwp_elig_membs_test_unlawful_conduct(elig_memb_count), 		6, 15, 43

				dwp_elig_membs_test_absence(elig_memb_count) = trim(dwp_elig_membs_test_absence(elig_memb_count))
				dwp_elig_membs_test_child_age(elig_memb_count) = trim(dwp_elig_membs_test_child_age(elig_memb_count))
				dwp_elig_membs_test_citizenship(elig_memb_count) = trim(dwp_elig_membs_test_citizenship(elig_memb_count))
				dwp_elig_membs_test_citizenship_verif(elig_memb_count) = trim(dwp_elig_membs_test_citizenship_verif(elig_memb_count))
				dwp_elig_membs_test_dupl_assistance(elig_memb_count) = trim(dwp_elig_membs_test_dupl_assistance(elig_memb_count))
				dwp_elig_membs_test_foster_care(elig_memb_count) = trim(dwp_elig_membs_test_foster_care(elig_memb_count))
				dwp_elig_membs_test_fraud(elig_memb_count) = trim(dwp_elig_membs_test_fraud(elig_memb_count))

				dwp_elig_membs_test_minor_living_arrangement(elig_memb_count) = trim(dwp_elig_membs_test_minor_living_arrangement(elig_memb_count))
				dwp_elig_membs_test_post_60_removal(elig_memb_count) = trim(dwp_elig_membs_test_post_60_removal(elig_memb_count))
				dwp_elig_membs_test_ssi(elig_memb_count) = trim(dwp_elig_membs_test_ssi(elig_memb_count))
				dwp_elig_membs_test_ssn_coop(elig_memb_count) = trim(dwp_elig_membs_test_ssn_coop(elig_memb_count))
				dwp_elig_membs_test_unit_member(elig_memb_count) = trim(dwp_elig_membs_test_unit_member(elig_memb_count))
				dwp_elig_membs_test_unlawful_conduct(elig_memb_count) = trim(dwp_elig_membs_test_unlawful_conduct(elig_memb_count))

				transmit

				Call write_value_and_transmit("X", row, 69)		'open member EMPS information
				EMReadScreen emps_exists_for_memb, 19, 24, 2
				If emps_exists_for_memb = "EMPS DOES NOT EXIST" Then
					EMWriteScreen " ", row, 69
				Else
					EMReadScreen dwp_elig_membs_es_status_code(elig_memb_count), 2, 9, 22
					EMReadScreen dwp_elig_membs_es_status_info(elig_memb_count), 30, 9, 25

					dwp_elig_membs_es_status_code(elig_memb_count) = trim(dwp_elig_membs_es_status_code(elig_memb_count))
					dwp_elig_membs_es_status_info(elig_memb_count) = trim(dwp_elig_membs_es_status_info(elig_memb_count))
					transmit
				End If

				row = row + 1
				elig_memb_count = elig_memb_count + 1
				EMReadScreen next_ref_numb, 2, row, 6
			Loop until next_ref_numb = "  "

			transmit 		'going to the next panel - DWCR

			EMReadScreen dwp_elig_case_test_application_withdrawn, 	6, 6, 7
			EMReadScreen dwp_elig_case_test_assets, 				6, 7, 7
			EMReadScreen dwp_elig_case_test_CS_disqualification, 	6, 8, 7
			EMReadScreen dwp_elig_case_test_death_of_applicant, 	6, 9, 7
			EMReadScreen dwp_elig_case_test_dupl_assistance, 		6, 10, 7
			EMReadScreen dwp_elig_case_test_eligible_child, 		6, 11, 7
			EMReadScreen dwp_elig_case_test_ES_disqualification, 	6, 12, 7
			EMReadScreen dwp_elig_case_test_fail_coop, 				6, 13, 7
			EMReadScreen dwp_elig_case_test_four_month_limit, 		6, 14, 7

			EMReadScreen dwp_elig_case_test_initial_income, 		6, 6, 45
			EMReadScreen dwp_elig_case_test_MFIP_conversion, 		6, 7, 45
			EMReadScreen dwp_elig_case_test_residence, 				6, 8, 45
			EMReadScreen dwp_elig_case_test_strike, 				6, 9, 45
			EMReadScreen dwp_elig_case_test_TANF_time_limit, 		6, 10, 45
			EMReadScreen dwp_elig_case_test_transfer_of_assets, 	6, 11, 45
			EMReadScreen dwp_elig_case_test_verif, 					6, 12, 45

			EMReadScreen dwp_elig_case_test_new_spouse_income, 		6, 17, 7

			dwp_elig_case_test_application_withdrawn = trim(dwp_elig_case_test_application_withdrawn)
			dwp_elig_case_test_assets = trim(dwp_elig_case_test_assets)
			dwp_elig_case_test_CS_disqualification = trim(dwp_elig_case_test_CS_disqualification)
			dwp_elig_case_test_death_of_applicant = trim(dwp_elig_case_test_death_of_applicant)
			dwp_elig_case_test_dupl_assistance = trim(dwp_elig_case_test_dupl_assistance)
			dwp_elig_case_test_eligible_child = trim(dwp_elig_case_test_eligible_child)
			dwp_elig_case_test_ES_disqualification = trim(dwp_elig_case_test_ES_disqualification)
			dwp_elig_case_test_fail_coop = trim(dwp_elig_case_test_fail_coop)
			dwp_elig_case_test_four_month_limit = trim(dwp_elig_case_test_four_month_limit)

			dwp_elig_case_test_initial_income = trim(dwp_elig_case_test_initial_income)
			dwp_elig_case_test_MFIP_conversion = trim(dwp_elig_case_test_MFIP_conversion)
			dwp_elig_case_test_residence = trim(dwp_elig_case_test_residence)
			dwp_elig_case_test_strike = trim(dwp_elig_case_test_strike)
			dwp_elig_case_test_TANF_time_limit = trim(dwp_elig_case_test_TANF_time_limit)
			dwp_elig_case_test_transfer_of_assets = trim(dwp_elig_case_test_transfer_of_assets)
			dwp_elig_case_test_verif = trim(dwp_elig_case_test_verif)

			dwp_elig_case_test_new_spouse_income = trim(dwp_elig_case_test_new_spouse_income)

			If dwp_elig_case_test_assets <> "NA" Then
				Call write_value_and_transmit("X", 7, 5)

				EMReadScreen dwp_elig_asset_CASH, 9, 8, 54
				EMReadScreen dwp_elig_asset_ACCT, 9, 9, 54
				EMReadScreen dwp_elig_asset_SECU, 9, 10, 54
				EMReadScreen dwp_elig_asset_CARS, 9, 11, 54
				EMReadScreen dwp_elig_asset_SPON, 9, 12, 54

				EMReadScreen dwp_elig_asset_total, 9, 17, 54
				EMReadScreen dwp_elig_asset_maximum, 9, 18, 54

				dwp_elig_asset_CASH = trim(dwp_elig_asset_CASH)
				dwp_elig_asset_ACCT = trim(dwp_elig_asset_ACCT)
				dwp_elig_asset_SECU = trim(dwp_elig_asset_SECU)
				dwp_elig_asset_CARS = trim(dwp_elig_asset_CARS)
				dwp_elig_asset_SPON = trim(dwp_elig_asset_SPON)
				dwp_elig_asset_total = trim(dwp_elig_asset_total)
				dwp_elig_asset_maximum = trim(dwp_elig_asset_maximum)

				transmit
			End If

			If dwp_elig_case_test_fail_coop <> "NA" Then
				Call write_value_and_transmit("X", 13, 5)

				EMReadScreen dwp_elig_test_fail_coop_applied_other_benefits, 6, 10, 30
				EMReadScreen dwp_elig_test_fail_coop_provide_requested_info, 6, 11, 30
				EMReadScreen dwp_elig_test_fail_coop_IEVS, 6, 12, 30
				EMReadScreen dwp_elig_test_fail_coop_vendor_info, 6, 13, 30

				dwp_elig_test_fail_coop_applied_other_benefits = trim(dwp_elig_test_fail_coop_applied_other_benefits)
				dwp_elig_test_fail_coop_provide_requested_info = trim(dwp_elig_test_fail_coop_provide_requested_info)
				dwp_elig_test_fail_coop_IEVS = trim(dwp_elig_test_fail_coop_IEVS)
				dwp_elig_test_fail_coop_vendor_info = trim(dwp_elig_test_fail_coop_vendor_info)

				transmit

			End If

			If dwp_elig_case_test_initial_income <> "NA" Then
				Call write_value_and_transmit("X", 6, 43)

				EMReadScreen dwp_elig_initial_counted_earned_income, 	9, 8, 42
				EMReadScreen dwp_elig_initial_dependent_care_expense, 	9, 9, 42
				EMReadScreen dwp_elig_initial_counted_unearned_incom, 	9, 10, 42
				EMReadScreen dwp_elig_initial_counted_deemed_income, 	9, 11, 42
				EMReadScreen dwp_elig_initial_child_support_exclusion, 	9, 12, 42
				EMReadScreen dwp_elig_initial_total_counted_income, 	9, 13, 42
				EMReadScreen dwp_elig_initial_family_wage_level, 		9, 15, 42

				dwp_elig_initial_counted_earned_income = trim(dwp_elig_initial_counted_earned_income)
				dwp_elig_initial_dependent_care_expense = trim(dwp_elig_initial_dependent_care_expense)
				dwp_elig_initial_counted_unearned_incom = trim(dwp_elig_initial_counted_unearned_incom)
				dwp_elig_initial_counted_deemed_income = trim(dwp_elig_initial_counted_deemed_income)
				dwp_elig_initial_child_support_exclusion = trim(dwp_elig_initial_child_support_exclusion)
				dwp_elig_initial_total_counted_income = trim(dwp_elig_initial_total_counted_income)
				dwp_elig_initial_family_wage_level = trim(dwp_elig_initial_family_wage_level)

				'TODO - read member specific detail'

				transmit
			End If

			If dwp_elig_case_test_verif <> "NA" Then
				Call write_value_and_transmit("X", 12, 43)

				EMReadScreen dwp_elig_test_verif_ACCT, 		6, 5, 32
				EMReadScreen dwp_elig_test_verif_BUSI, 		6, 6, 32
				EMReadScreen dwp_elig_test_verif_CARS, 		6, 7, 32
				EMReadScreen dwp_elig_test_verif_JOBS, 		6, 8, 32
				EMReadScreen dwp_elig_test_verif_MEMB_dob, 	6, 9, 32
				EMReadScreen dwp_elig_test_verif_MEMB_id, 	6, 10, 32
				EMReadScreen dwp_elig_test_verif_PARE, 		6, 11, 32
				EMReadScreen dwp_elig_test_verif_PREG, 		6, 12, 32
				EMReadScreen dwp_elig_test_verif_RBIC, 		6, 13, 32
				EMReadScreen dwp_elig_test_verif_ADDR, 		6, 14, 32
				EMReadScreen dwp_elig_test_verif_SCHL, 		6, 15, 32
				EMReadScreen dwp_elig_test_verif_SECU, 		6, 16, 32
				EMReadScreen dwp_elig_test_verif_SPON, 		6, 17, 32
				EMReadScreen dwp_elig_test_verif_UNEA, 		6, 18, 32

				dwp_elig_test_verif_ACCT = trim(dwp_elig_test_verif_ACCT)
				dwp_elig_test_verif_BUSI = trim(dwp_elig_test_verif_BUSI)
				dwp_elig_test_verif_CARS = trim(dwp_elig_test_verif_CARS)
				dwp_elig_test_verif_JOBS = trim(dwp_elig_test_verif_JOBS)
				dwp_elig_test_verif_MEMB_dob = trim(dwp_elig_test_verif_MEMB_dob)
				dwp_elig_test_verif_MEMB_id = trim(dwp_elig_test_verif_MEMB_id)
				dwp_elig_test_verif_PARE = trim(dwp_elig_test_verif_PARE)
				dwp_elig_test_verif_PREG = trim(dwp_elig_test_verif_PREG)
				dwp_elig_test_verif_RBIC = trim(dwp_elig_test_verif_RBIC)
				dwp_elig_test_verif_ADDR = trim(dwp_elig_test_verif_ADDR)
				dwp_elig_test_verif_SCHL = trim(dwp_elig_test_verif_SCHL)
				dwp_elig_test_verif_SECU = trim(dwp_elig_test_verif_SECU)
				dwp_elig_test_verif_SPON = trim(dwp_elig_test_verif_SPON)
				dwp_elig_test_verif_UNEA = trim(dwp_elig_test_verif_UNEA)

				transmit
			End If

			If dwp_elig_case_test_new_spouse_income <> "NA" Then
				Call write_value_and_transmit("X", 17, 5)

				'TODO - Read New Spouse Income Information

				transmit
			End If

			transmit 		'going to the next panel - DWCB1


			EMReadScreen dwp_elig_budg_shel_rent_mortgage, 		9, 5, 29
			EMReadScreen dwp_elig_budg_shel_property_tax, 		9, 6, 29
			EMReadScreen dwp_elig_budg_shel_house_insurance, 	9, 7, 29
			EMReadScreen dwp_elig_budg_hest_electricity, 		9, 8, 29
			EMReadScreen dwp_elig_budg_hest_heat_air, 			9, 9, 29
			EMReadScreen dwp_elig_budg_hest_water_sewer_garbage, 9, 10, 29
			EMReadScreen dwp_elig_budg_hest_phone, 				9, 11, 29
			EMReadScreen dwp_elig_budg_shel_other, 				9, 12, 29

			EMReadScreen dwp_elig_budg_total_shelter_costs, 	9, 14, 29
			EMReadScreen dwp_elig_budg_personal_needs, 			9, 15, 29

			EMReadScreen dwp_elig_budg_total_DWP_need, 			9, 17, 29

			EMReadScreen dwp_elig_budg_earned_income, 			9, 7, 71
			EMReadScreen dwp_elig_budg_unearned_income, 		9, 8, 71
			EMReadScreen dwp_elig_budg_deemed_income, 			9, 9, 71
			EMReadScreen dwp_elig_budg_child_support_exclusion, 9, 10, 71
			EMReadScreen dwp_elig_budg_budget_month_total, 		9, 11, 71
			EMReadScreen dwp_elig_budg_prior_low, 				9, 12, 71
			EMReadScreen dwp_elig_budg_DWP_countable_income, 	9, 13, 71

			EMReadScreen dwp_elig_budg_unmet_need, 				9, 15, 71
			EMReadScreen dwp_elig_budg_DWP_max_grant, 			9, 16, 71
			EMReadScreen dwp_elig_budg_DWP_grant, 				9, 17, 71

			dwp_elig_budg_shel_rent_mortgage = trim(dwp_elig_budg_shel_rent_mortgage)
			dwp_elig_budg_shel_property_tax = trim(dwp_elig_budg_shel_property_tax)
			dwp_elig_budg_shel_house_insurance = trim(dwp_elig_budg_shel_house_insurance)
			dwp_elig_budg_hest_electricity = trim(dwp_elig_budg_hest_electricity)
			dwp_elig_budg_hest_heat_air = trim(dwp_elig_budg_hest_heat_air)
			dwp_elig_budg_hest_water_sewer_garbage = trim(dwp_elig_budg_hest_water_sewer_garbage)
			dwp_elig_budg_hest_phone = trim(dwp_elig_budg_hest_phone)
			dwp_elig_budg_shel_other = trim(dwp_elig_budg_shel_other)
			dwp_elig_budg_total_shelter_costs = trim(dwp_elig_budg_total_shelter_costs)
			dwp_elig_budg_personal_needs = trim(dwp_elig_budg_personal_needs)
			dwp_elig_budg_total_DWP_need = trim(dwp_elig_budg_total_DWP_need)
			dwp_elig_budg_earned_income = trim(dwp_elig_budg_earned_income)
			dwp_elig_budg_unearned_income = trim(dwp_elig_budg_unearned_income)
			dwp_elig_budg_deemed_income = trim(dwp_elig_budg_deemed_income)
			dwp_elig_budg_child_support_exclusion = trim(dwp_elig_budg_child_support_exclusion)
			dwp_elig_budg_budget_month_total = trim(dwp_elig_budg_budget_month_total)
			dwp_elig_budg_prior_low = trim(dwp_elig_budg_prior_low)
			dwp_elig_budg_DWP_countable_income = trim(dwp_elig_budg_DWP_countable_income)
			dwp_elig_budg_unmet_need = trim(dwp_elig_budg_unmet_need)
			dwp_elig_budg_DWP_max_grant = trim(dwp_elig_budg_DWP_max_grant)
			dwp_elig_budg_DWP_grant = trim(dwp_elig_budg_DWP_grant)

			Call write_value_and_transmit("X", 7, 41)
			EmReadScreen pop_up_menu_title, 13, 3, 46
			If pop_up_menu_title = "Earned Income" Then
				'TODO - read member specific unearned income
				transmit
			End If

			Call write_value_and_transmit("X", 8, 41)
			EmReadScreen pop_up_menu_title, 15, 5, 32
			If pop_up_menu_title = "Unearned Income" Then
				'TODO - read member specific unearned income
				transmit
			End If

			Call write_value_and_transmit("X", 9, 41)
			EmReadScreen pop_up_menu_title, 13, 3, 36
			If pop_up_menu_title = "Deemed Income" Then
				'TODO - read member specific unearned income
				' EMReadScreen dwp_elig_membs_budg_deemed_self_emp(member_sel), 				9, 8, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_wages(member_sel), 					9, 9, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_counted_earned(member_sel), 		9, 10, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_standard_EI_disregard(member_sel), 	9, 11, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_earned_subtotal(member_sel), 		9, 12, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_earned_disregard(member_sel), 		9, 13, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_unearned_income(member_sel), 		9, 14, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_subtotal_counted_income(member_sel), 9, 15, 56
				'
				' EMReadScreen dwp_elig_membs_budg_deemed_deemer_unmet_need(member_sel), 		9, 18, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_allocation(member_sel), 			9, 19, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_child_support(member_sel), 			9, 20, 56
				' EMReadScreen dwp_elig_membs_budg_deemed_counted_income(member_sel), 		9, 21, 56
				transmit
			End If

			Call write_value_and_transmit("X", 10, 41)
			EMReadScreen dwp_elig_cses_income, 9, 10, 54
			EMReadScreen dwp_elig_child_count, 2, 12, 36
			'TODO - read member specific unearned income

			dwp_elig_cses_income = trim(dwp_elig_cses_income)
			dwp_elig_child_count = trim(dwp_elig_child_count)

			transmit


			transmit 		'going to the next panel - DWB2

			EMReadScreen dwp_elig_prorated_date, 8, 6, 18
			If dwp_elig_prorated_date = "__ __ __" then dwp_elig_prorated_date = ""
			dwp_elig_prorated_date = replace(dwp_elig_prorated_date, " ", "/")

			EMReadScreen dwp_elig_prorated_amount, 9, 6, 35

			EMReadScreen dwp_elig_amount_already_issued, 	9, 9, 35
			EMReadScreen dwp_elig_supplement_due, 			9, 10, 35
			EMReadScreen dwp_elig_overpayment, 				9, 11, 35
			EMReadScreen dwp_elig_adjusted_grant_amount, 	9, 12, 35
			EMReadScreen dwp_elig_recoupment_amount, 		9, 13, 35

			EMReadScreen dwp_elig_shelter_benefit_grant, 	9, 15, 35
			EMReadScreen dwp_elig_personal_needs_grant, 	9, 16, 35

			Call write_value_and_transmit("X", 11, 3)
			EMReadScreen dwp_elig_overpayment_fed_hh_count, 	2, 10, 31
			EMReadScreen dwp_elig_overpayment_fed_amount, 		9, 10, 50
			EMReadScreen dwp_elig_overpayment_state_hh_count, 	2, 12, 31
			EMReadScreen dwp_elig_overpayment_state_amount, 	9, 12, 50
			transmit

			Call write_value_and_transmit("X", 12, 3)
			EMReadScreen dwp_elig_adjusted_grant_fed_hh_count, 		2, 10, 25
			EMReadScreen dwp_elig_adjusted_grant_fed_amount, 		9, 10, 45
			EMReadScreen dwp_elig_adjusted_grant_state_hh_count, 	2, 12, 25
			EMReadScreen dwp_elig_adjusted_grant_state_amount, 		9, 12, 45
			transmit

			dwp_elig_prorated_amount = trim(dwp_elig_prorated_amount)
			dwp_elig_amount_already_issued = trim(dwp_elig_amount_already_issued)
			dwp_elig_supplement_due = trim(dwp_elig_supplement_due)
			dwp_elig_overpayment = trim(dwp_elig_overpayment)
			dwp_elig_adjusted_grant_amount = trim(dwp_elig_adjusted_grant_amount)
			dwp_elig_recoupment_amount = trim(dwp_elig_recoupment_amount)
			dwp_elig_shelter_benefit_grant = trim(dwp_elig_shelter_benefit_grant)
			dwp_elig_personal_needs_grant = trim(dwp_elig_personal_needs_grant)
			dwp_elig_overpayment_fed_amount = trim(dwp_elig_overpayment_fed_amount)
			dwp_elig_overpayment_state_amount = trim(dwp_elig_overpayment_state_amount)
			dwp_elig_adjusted_grant_fed_amount = trim(dwp_elig_adjusted_grant_fed_amount)
			dwp_elig_adjusted_grant_state_amount = trim(dwp_elig_adjusted_grant_state_amount)

			transmit 		'going to the next panel - DWSM

			EMReadScreen dwp_approved_date, 8, 3, 14
			EMReadScreen dwp_process_date, 8, 2, 73
			EMReadScreen dwp_prev_approval, 4, 3, 73

			EMReadScreen dwp_case_last_approval_date, 8, 5, 31
			EMReadScreen dwp_case_current_prog_status, 12, 6, 31
			EMReadScreen dwp_case_eligibility_result, 12,  7, 31
			EMReadScreen dwp_case_source_of_info, 4, 9, 31
			EMReadScreen dwp_case_benefit_impact, 12, 10, 31
			EMReadScreen dwp_case_4th_month_of_elig, 5, 11, 31
			EMReadScreen dwp_case_caregivers_have_es_plan, 1, 12, 31
			EMReadScreen dwp_case_responsible_county, 2, 13, 31
			EMReadScreen dwp_case_service_county, 2, 14, 31

			EMReadScreen dwp_case_asst_unit_caregivers, 3, 5, 72
			EMReadScreen dwp_case_asst_unit_children, 3, 6, 72
			EMReadScreen dwp_case_total_assets, 10, 7, 71
			EMReadScreen dwp_case_maximum_assets, 10, 8, 71
			EMReadScreen dwp_case_summary_grant_amount, 10, 10, 71
			EMReadScreen dwp_case_summary_net_grant_amount, 10, 12, 71
			EMReadScreen dwp_case_summary_shelter_benefit_portion, 10, 13, 71
			EMReadScreen dwp_case_summary_personal_needs_portion, 10, 14, 71

			dwp_prev_approval = trim(dwp_prev_approval)
			dwp_case_last_approval_date = trim(dwp_case_last_approval_date)

			dwp_case_current_prog_status = trim(dwp_case_current_prog_status)
			dwp_case_eligibility_result = trim(dwp_case_eligibility_result)
			dwp_case_source_of_info = trim(dwp_case_source_of_info)
			dwp_case_benefit_impact = trim(dwp_case_benefit_impact)

			dwp_case_asst_unit_caregivers = trim(dwp_case_asst_unit_caregivers)
			dwp_case_asst_unit_children = trim(dwp_case_asst_unit_children)
			dwp_case_total_assets = trim(dwp_case_total_assets)
			dwp_case_maximum_assets = trim(dwp_case_maximum_assets)
			dwp_case_summary_grant_amount = trim(dwp_case_summary_grant_amount)
			dwp_case_summary_net_grant_amount = trim(dwp_case_summary_net_grant_amount)
			dwp_case_summary_shelter_benefit_portion = trim(dwp_case_summary_shelter_benefit_portion)
			dwp_case_summary_personal_needs_portion = trim(dwp_case_summary_personal_needs_portion)
		End If
		Call back_to_SELF
	end sub
end class

class mfip_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found
	public er_month
	public hrf_month
	public er_status
	public er_caf_date
	public er_interview_date
	public hrf_status
	public hrf_doc_date

	public mfip_elig_ref_numbs()
	public mfip_elig_membs_full_name()
	public mfip_elig_membs_request_yn()
	public mfip_elig_membs_code()
	public mfip_elig_membs_status_info()
	public mfip_elig_membs_deemed()
	public mfip_elig_membs_counted()
	public mfip_elig_membs_eligibility()
	public mfip_elig_membs_begin_date()
	public mfip_elig_membs_budget_cycle()
	public mfip_elig_membs_absence()
	public mfip_elig_membs_child_age()
	public mfip_elig_membs_citizenship()
	public mfip_elig_membs_citizenship_verif()
	public mfip_elig_membs_dupl_assist()
	public mfip_elig_membs_foster_care()
	public mfip_elig_membs_fraud()
	public mfip_elig_membs_fs_disq()
	public mfip_elig_membs_minor_living_arngmt()
	public mfip_elig_membs_post_60_removal()
	public mfip_elig_membs_ssi()
	public mfip_elig_membs_ssn_code()
	public mfip_elig_membs_unit_memb()
	public mfip_elig_membs_unlawful_conduct()
	public mfip_elig_membs_fs_recvd()
	public mfip_elig_membs_es_status_code()
	public mfip_elig_membs_es_status_info()

	public mfip_memb_cash_portion_code()
	public mfip_memb_food_portion_code()
	public mfip_memb_state_food_code()
	public mfip_memb_sanction_yn()
	public mfip_memb_sanction_child_support_test()
	public mfip_memb_sanction_drug_felon_test()
	public mfip_memb_sanction_emp_services_test()
	public mfip_memb_sanction_fin_orient_test()
	public mfip_memb_sanction_occurence()
	public mfip_memb_sanction_begin_date()
	public mfip_memb_sanction_last_sanc_month()

	public mfip_elig_membs_initial_BUSI_inc_total()
	public mfip_elig_membs_initial_JOBS_inc_total()
	public mfip_elig_membs_initial_earned_inc_total()
	public mfip_elig_membs_initial_stndrd_ei_disregard()
	public mfip_elig_membs_initial_earned_inc_subtotal()
	public mfip_elig_membs_initial_earned_inc_disregard()
	public mfip_elig_membs_initial_avail_earned_inc()
	public mfip_elig_membs_initial_allocation()
	public mfip_elig_membs_initial_child_support()
	public mfip_elig_membs_initial_counted_earned_inc_total()
	public mfip_elig_membs_initial_UNEA_inc_total()
	public mfip_elig_membs_initial_allocation_balance()
	public mfip_elig_membs_initial_child_support_balance()
	public mfip_elig_membs_initial_counted_UNEA_inc_total()
	public mfip_elig_membs_initial_income_cses_retro_income()
	public mfip_elig_membs_initial_income_cses_prosp_income()
	public mfip_elig_membs_new_spouse_earned_income()
	public mfip_elig_membs_new_spouse_unearned_income()
	public mfip_elig_membs_new_spouse_total_income()

	public mfip_elig_membs_self_emp_income()
	public mfip_elig_membs_wages_income()
	public mfip_elig_membs_total_earned_income()
	public mfip_elig_membs_standard_EI_disregard()
	public mfip_elig_membs_earned_income_subtotal()
	public mfip_elig_membs_earned_income_50_perc_disregard()
	public mfip_elig_membs_available_earned_income()
	public mfip_elig_membs_allocation_deduction()
	public mfip_elig_membs_child_support_deduction()
	public mfip_elig_membs_counted_earned_income()

	public mfip_elig_membs_total_unearned_income()
	public mfip_elig_membs_allocation_balance()
	public mfip_elig_membs_child_support_balance()
	public mfip_elig_membs_counted_unearned_income()

	public mfip_elig_membs_county_88_cses_income()
	public mfip_elig_membs_county_88_gaming_income()
	public mfip_elig_membs_county_88_200_perc_fpg()
	public mfip_elig_membs_county_88_deemers_unmet_need()
	public mfip_elig_membs_county_88_allocation()
	public mfip_elig_membs_county_88_child_support()
	public mfip_elig_membs_county_88_counted_gaming_income()

	public mfip_elig_membs_retro_subsidy_amount()
	public mfip_elig_membs_prosp_subsidy_amount()

	public mfip_cash_opt_out
	public mfip_HG_opt_out
	public mfip_child_only
	public mfip_case_in_sancttion

	public mfip_case_test_appl_withdraw
	public mfip_case_test_asset
	public mfip_case_test_death_applicant
	public mfip_case_test_dupl_assist
	public mfip_case_test_elig_child
	public mfip_case_test_fail_coop
	public mfip_case_test_fail_file
	public mfip_case_test_initial_income
	public mfip_case_test_minor_liv_arrange
	public mfip_case_test_monthly_income
	public mfip_case_test_post_60_disq
	public mfip_case_test_residence
	public mfip_case_test_sanction_limit
	public mfip_case_test_strike
	public mfip_case_test_TANF_time_limit
	public mfip_case_test_transfer_asset
	public mfip_case_test_verif
	public mfip_case_test_275_new_spouse_income
	public mfip_fs_case_test_fail_coop_snap_qc
	public mfip_fs_case_test_opt_out_cash
	public mfip_fs_case_test_opt_out_housing_grant

	public mfip_counted_asset_CASH
	public mfip_counted_asset_ACCT
	public mfip_counted_asset_SECU
	public mfip_counted_asset_CARS
	public mfip_counted_asset_SPON
	public mfip_counted_asset_total
	public mfip_counted_asset_max

	public mfip_initial_income_earned
	public mfip_initial_income_deoendant_care
	public mfip_initial_income_unearned
	public mfip_initial_income_deemed
	public mfip_initial_income_cses_exclusion
	public mfip_initial_income_cses_income
	public mfip_initial_income_cses_child_count
	public mfip_initial_income_net_cses_income
	public mfip_initial_income_total
	public mfip_initial_income_family_wage_level

	public mfip_12_month_start_date
	public mfip_designated_spouse_ref_numb
	public mfip_new_spouse_inc_earned
	public mfip_new_spouse_inc_unearned
	public mfip_new_spouse_inc_deemed_earned
	public mfip_new_spouse_inc_deemed_unearned
	public mfip_new_spouse_inc_total
	public mfip_275_fpg_amt
	public mfip_hh_size_fornew_spouse_calc

	public mfip_case_sanction_percent
	public mfip_case_sanction_vendor_yn
	public mfip_case_sanction_last_vendor_month

	public mfip_case_budg_family_wage_level
	public mfip_case_budg_monthly_earned_income
	public mfip_case_budg_wage_level_earned_inc_difference
	public mfip_case_budg_transitional_standard
	public mfip_case_budg_monthly_need
	public mfip_case_budg_unearned_income
	public mfip_case_budg_deemed_income
	public mfip_case_budg_cses_exclusion
	public mfip_case_budg_unmet_need
	public mfip_case_budg_unmet_need_food_potion
	public mfip_case_budg_tribal_counted_income
	public mfip_case_budg_unmet_need_cash_portion
	public mfip_case_budg_deduction_subsidy_tribal_cses

	public mfip_case_budg_net_food_portion
	public mfip_case_budg_net_cash_portion
	public mfip_case_budg_net_unmet_need
	public mfip_case_budg_deduction_sanction_vendor
	public mfip_case_budg_unmet_neet_subtotal
	public mfip_case_budg_subtotal_food_portion
	public mfip_case_budg_food_portion_deduction
	public mfip_case_budg_entitlement_food_portion
	public mfip_case_budg_entitlement_housing_grant

	public mfip_budg_cses_excln_cses_income
	public mfip_budg_cses_excln_child_count
	public mfip_budg_cses_excln_total
	public mfip_budg_total_county_88_child_support_income
	public mfip_budg_total_county_88_gaming_income
	public mfip_budg_total_tribal_income_fs_portion_deduction
	public mfip_budg_total_housing_subsidy_amount
	public mfip_budg_total_tribal_child_support
	public mfip_budg_total_subsidy_tribal_cash_portion_deduction
	public mfip_elig_budg_total_countable_housing_subsidy
	public mfip_elig_budg_housing_subsidy_exempt

	public mfip_case_budg_10_perc_sanc
	public mfip_case_budg_unmet_need_after_pre_vndr_sanc
	public mfip_case_budg_sanc_calc_food_portion
	public mfip_case_budg_sanc_calc_cash_portion
	public mfip_case_budg_pot_mand_vndr_pymts
	public mfip_case_budg_30_perc_sanc

	public mfip_case_budg_non_citzn_fs_inelig_pers_count
	public mfip_case_budg_non_citzn_fs_inelig_amt
	public mfip_case_budg_other_fs_inelig_pers_count
	public mfip_case_budg_other_fs_inelig_amt

	public mfip_case_budg_prorate_date
	public mfip_case_budg_fed_food_benefit
	public mfip_case_budg_food_prorated_amt
	public mfip_case_budg_entitlement_cash_portion
	public mfip_case_budg_mand_sanc_vendor
	public mfip_case_budg_net_cash_after_sanc_portion
	public mfip_case_budg_cash_prorated_amt
	public mfip_case_budg_state_food_benefit
	public mfip_case_budg_state_food_prorated_amt
	public mfip_case_budg_grant_amount
	public mfip_case_budg_amt_already_issued
	public mfip_case_budg_supplement_due
	public mfip_case_budg_overpayment
	public mfip_case_budg_adjusted_grant_amt
	public mfip_case_budg_recoupment
	public mfip_case_budg_total_food_issuance
	public mfip_case_budg_total_cash_issuance
	public mfip_case_budg_total_housing_grant_issuance

	public mfip_case_budg_food_supplement
	public mfip_case_budg_state_food_supplement
	public mfip_case_budg_cash_supplement
	public mfip_case_budg_housing_grant_supplement

	public mfip_case_budg_cash_recoupment
	public mfip_case_budg_state_food_recoupment
	public mfip_case_budg_food_recoupment

	public mfip_case_budg_fed_food_memb_count
	public mfip_case_budg_fed_food_benefit_amt
	public mfip_case_budg_state_food_memb_count
	public mfip_case_budg_state_food_benefit_amt

	public mfip_case_budg_tanf_cash_memb_count
	public mfip_case_budg_tanf_cash_benefit_amt
	public mfip_case_budg_state_cash_memb_count
	public mfip_case_budg_state_cash_benefit_amt

	public mfip_approved_date
	public mfip_process_date
	public mfip_prev_approval
	public mfip_case_last_approval_date
	public mfip_case_current_prog_status
	public mfip_case_eligibility_result
	public mfip_case_hrf_reporting
	public mfip_case_source_of_info
	public mfip_case_benefit_impact
	public mfip_case_review_date
	public mfip_case_budget_cycle
	public mfip_case_vendor_reason_code
	public mfip_case_vendor_reason_info
	public mfip_case_responsible_county
	public mfip_case_service_county
	public mfip_case_asst_unit_caregivers
	public mfip_case_asst_unit_children
	public mfip_case_total_assets
	public mfip_case_maximum_assets
	public mfip_case_summary_grant_amount
	public mfip_case_summary_net_grant_amount
	public mfip_case_summary_cash_portion
	public mfip_case_summary_food_portion
	public mfip_case_summary_housing_grant

	public sub read_elig()
		mfip_cash_opt_out = False
		mfip_HG_opt_out = False
		mfip_child_only = False
		mfip_case_in_sancttion = False

		approved_today = False
		approved_version_found = False

		call navigate_to_MAXIS_screen("ELIG", "MFIP")
		EMWriteScreen elig_footer_month, 20, 55
		EMWriteScreen elig_footer_year, 20, 58
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

			ReDim mfip_elig_ref_numbs(0)
			ReDim mfip_elig_membs_full_name(0)
			ReDim mfip_elig_membs_request_yn(0)
			ReDim mfip_elig_membs_code(0)
			ReDim mfip_elig_membs_status_info(0)
			ReDim mfip_elig_membs_deemed(0)
			ReDim mfip_elig_membs_counted(0)
			ReDim mfip_elig_membs_eligibility(0)
			ReDim mfip_elig_membs_begin_date(0)
			ReDim mfip_elig_membs_budget_cycle(0)
			ReDim mfip_elig_membs_absence(0)
			ReDim mfip_elig_membs_child_age(0)
			ReDim mfip_elig_membs_citizenship(0)
			ReDim mfip_elig_membs_citizenship_verif(0)
			ReDim mfip_elig_membs_dupl_assist(0)
			ReDim mfip_elig_membs_foster_care(0)
			ReDim mfip_elig_membs_fraud(0)
			ReDim mfip_elig_membs_fs_disq(0)
			ReDim mfip_elig_membs_minor_living_arngmt(0)
			ReDim mfip_elig_membs_post_60_removal(0)
			ReDim mfip_elig_membs_ssi(0)
			ReDim mfip_elig_membs_ssn_code(0)
			ReDim mfip_elig_membs_unit_memb(0)
			ReDim mfip_elig_membs_unlawful_conduct(0)
			ReDim mfip_elig_membs_fs_recvd(0)
			ReDim mfip_elig_membs_es_status_code(0)
			ReDim mfip_elig_membs_es_status_info(0)
			ReDim mfip_memb_cash_portion_code(0)
			ReDim mfip_memb_food_portion_code(0)
			ReDim mfip_memb_state_food_code(0)
			ReDim mfip_memb_sanction_yn(0)
			ReDim mfip_memb_sanction_child_support_test(0)
			ReDim mfip_memb_sanction_drug_felon_test(0)
			ReDim mfip_memb_sanction_emp_services_test(0)
			ReDim mfip_memb_sanction_fin_orient_test(0)
			ReDim mfip_memb_sanction_occurence(0)
			ReDim mfip_memb_sanction_begin_date(0)
			ReDim mfip_memb_sanction_last_sanc_month(0)
			ReDim mfip_elig_membs_initial_BUSI_inc_total(0)
			ReDim mfip_elig_membs_initial_JOBS_inc_total(0)
			ReDim mfip_elig_membs_initial_earned_inc_total(0)
			ReDim mfip_elig_membs_initial_stndrd_ei_disregard(0)
			ReDim mfip_elig_membs_initial_earned_inc_subtotal(0)
			ReDim mfip_elig_membs_initial_earned_inc_disregard(0)
			ReDim mfip_elig_membs_initial_avail_earned_inc(0)
			ReDim mfip_elig_membs_initial_allocation(0)
			ReDim mfip_elig_membs_initial_child_support(0)
			ReDim mfip_elig_membs_initial_counted_earned_inc_total(0)
			ReDim mfip_elig_membs_initial_UNEA_inc_total(0)
			ReDim mfip_elig_membs_initial_allocation_balance(0)
			ReDim mfip_elig_membs_initial_child_support_balance(0)
			ReDim mfip_elig_membs_initial_counted_UNEA_inc_total(0)
			ReDim mfip_elig_membs_initial_income_cses_retro_income(0)
			ReDim mfip_elig_membs_initial_income_cses_prosp_income(0)
			ReDim mfip_elig_membs_new_spouse_earned_income(0)
			ReDim mfip_elig_membs_new_spouse_unearned_income(0)
			ReDim mfip_elig_membs_new_spouse_total_income(0)
			ReDim mfip_elig_membs_self_emp_income(0)
			ReDim mfip_elig_membs_wages_income(0)
			ReDim mfip_elig_membs_total_earned_income(0)
			ReDim mfip_elig_membs_standard_EI_disregard(0)
			ReDim mfip_elig_membs_earned_income_subtotal(0)
			ReDim mfip_elig_membs_earned_income_50_perc_disregard(0)
			ReDim mfip_elig_membs_available_earned_income(0)
			ReDim mfip_elig_membs_allocation_deduction(0)
			ReDim mfip_elig_membs_child_support_deduction(0)
			ReDim mfip_elig_membs_counted_earned_income(0)
			ReDim mfip_elig_membs_total_unearned_income(0)
			ReDim mfip_elig_membs_allocation_balance(0)
			ReDim mfip_elig_membs_child_support_balance(0)
			ReDim mfip_elig_membs_counted_unearned_income(0)
			ReDim mfip_elig_membs_county_88_cses_income(0)
			ReDim mfip_elig_membs_county_88_gaming_income(0)
			ReDim mfip_elig_membs_county_88_200_perc_fpg(0)
			ReDim mfip_elig_membs_county_88_deemers_unmet_need(0)
			ReDim mfip_elig_membs_county_88_allocation(0)
			ReDim mfip_elig_membs_county_88_child_support(0)
			ReDim mfip_elig_membs_county_88_counted_gaming_income(0)
			ReDim mfip_elig_membs_retro_subsidy_amount(0)
			ReDim mfip_elig_membs_prosp_subsidy_amount(0)

			row = 7
			elig_memb_count = 0
			Do
				EMReadScreen ref_numb, 2, row, 6

				ReDim preserve mfip_elig_ref_numbs(elig_memb_count)
				ReDim preserve mfip_elig_membs_full_name(elig_memb_count)
				ReDim preserve mfip_elig_membs_request_yn(elig_memb_count)
				ReDim preserve mfip_elig_membs_code(elig_memb_count)
				ReDim preserve mfip_elig_membs_status_info(elig_memb_count)
				ReDim preserve mfip_elig_membs_deemed(elig_memb_count)
				ReDim preserve mfip_elig_membs_counted(elig_memb_count)
				ReDim preserve mfip_elig_membs_eligibility(elig_memb_count)
				ReDim preserve mfip_elig_membs_begin_date(elig_memb_count)
				ReDim preserve mfip_elig_membs_budget_cycle(elig_memb_count)
				ReDim preserve mfip_elig_membs_absence(elig_memb_count)
				ReDim preserve mfip_elig_membs_child_age(elig_memb_count)
				ReDim preserve mfip_elig_membs_citizenship(elig_memb_count)
				ReDim preserve mfip_elig_membs_citizenship_verif(elig_memb_count)
				ReDim preserve mfip_elig_membs_dupl_assist(elig_memb_count)
				ReDim preserve mfip_elig_membs_foster_care(elig_memb_count)
				ReDim preserve mfip_elig_membs_fraud(elig_memb_count)
				ReDim preserve mfip_elig_membs_fs_disq(elig_memb_count)
				ReDim preserve mfip_elig_membs_minor_living_arngmt(elig_memb_count)
				ReDim preserve mfip_elig_membs_post_60_removal(elig_memb_count)
				ReDim preserve mfip_elig_membs_ssi(elig_memb_count)
				ReDim preserve mfip_elig_membs_ssn_code(elig_memb_count)
				ReDim preserve mfip_elig_membs_unit_memb(elig_memb_count)
				ReDim preserve mfip_elig_membs_unlawful_conduct(elig_memb_count)
				ReDim preserve mfip_elig_membs_fs_recvd(elig_memb_count)
				ReDim preserve mfip_elig_membs_es_status_code(elig_memb_count)
				ReDim preserve mfip_elig_membs_es_status_info(elig_memb_count)
				ReDim preserve mfip_memb_cash_portion_code(elig_memb_count)
				ReDim preserve mfip_memb_food_portion_code(elig_memb_count)
				ReDim preserve mfip_memb_state_food_code(elig_memb_count)
				ReDim preserve mfip_memb_sanction_yn(elig_memb_count)
				ReDim preserve mfip_memb_sanction_child_support_test(elig_memb_count)
				ReDim preserve mfip_memb_sanction_drug_felon_test(elig_memb_count)
				ReDim preserve mfip_memb_sanction_emp_services_test(elig_memb_count)
				ReDim preserve mfip_memb_sanction_fin_orient_test(elig_memb_count)
				ReDim preserve mfip_memb_sanction_occurence(elig_memb_count)
				ReDim preserve mfip_memb_sanction_begin_date(elig_memb_count)
				ReDim preserve mfip_memb_sanction_last_sanc_month(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_BUSI_inc_total(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_JOBS_inc_total(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_earned_inc_total(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_stndrd_ei_disregard(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_earned_inc_subtotal(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_earned_inc_disregard(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_avail_earned_inc(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_allocation(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_child_support(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_counted_earned_inc_total(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_UNEA_inc_total(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_allocation_balance(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_child_support_balance(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_counted_UNEA_inc_total(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_income_cses_retro_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_initial_income_cses_prosp_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_new_spouse_earned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_new_spouse_unearned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_new_spouse_total_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_self_emp_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_wages_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_total_earned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_standard_EI_disregard(elig_memb_count)
				ReDim preserve mfip_elig_membs_earned_income_subtotal(elig_memb_count)
				ReDim preserve mfip_elig_membs_earned_income_50_perc_disregard(elig_memb_count)
				ReDim preserve mfip_elig_membs_available_earned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_allocation_deduction(elig_memb_count)
				ReDim preserve mfip_elig_membs_child_support_deduction(elig_memb_count)
				ReDim preserve mfip_elig_membs_counted_earned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_total_unearned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_allocation_balance(elig_memb_count)
				ReDim preserve mfip_elig_membs_child_support_balance(elig_memb_count)
				ReDim preserve mfip_elig_membs_counted_unearned_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_cses_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_gaming_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_200_perc_fpg(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_deemers_unmet_need(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_allocation(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_child_support(elig_memb_count)
				ReDim preserve mfip_elig_membs_county_88_counted_gaming_income(elig_memb_count)
				ReDim preserve mfip_elig_membs_retro_subsidy_amount(elig_memb_count)
				ReDim preserve mfip_elig_membs_prosp_subsidy_amount(elig_memb_count)

				mfip_elig_ref_numbs(elig_memb_count) = ref_numb
				EMReadScreen full_name_information, 20, row, 10
				full_name_information = trim(full_name_information)
				name_array = split(full_name_information, " ")
				For each name_parts in name_array
					If len(name_parts) <> 1 Then mfip_elig_membs_full_name(elig_memb_count) = mfip_elig_membs_full_name(elig_memb_count) & " " & name_parts
				Next
				mfip_elig_membs_full_name(elig_memb_count) = trim((mfip_elig_membs_full_name(elig_memb_count)))
				EMReadScreen mfip_elig_membs_request_yn(elig_memb_count), 1, row, 32
				EMReadScreen mfip_elig_membs_code(elig_memb_count), 1, row, 36
				EMReadScreen mfip_elig_membs_counted(elig_memb_count), 11, row, 41
				EMReadScreen mfip_elig_membs_eligibility(elig_memb_count), 10, row, 53
				EMReadScreen mfip_elig_membs_begin_date(elig_memb_count), 8, row, 67
				EMReadScreen mfip_elig_membs_budget_cycle(elig_memb_count), 1, row, 78

				If mfip_elig_membs_code(elig_memb_count) = "A" Then mfip_elig_membs_status_info(elig_memb_count) = "Eligible"
				If mfip_elig_membs_code(elig_memb_count) = "D" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed SSI, IV-E ADOPTION ASSISTANCE"
				If mfip_elig_membs_code(elig_memb_count) = "F" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed FRAUD, SSN COOP, UNLAWFUL CONDUCT"
				If mfip_elig_membs_code(elig_memb_count) = "G" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Parent of a minor caregiver"
				If mfip_elig_membs_code(elig_memb_count) = "H" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed CITIZENSHIP, CITIZENSHIP VERIFICATION"
				If mfip_elig_membs_code(elig_memb_count) = "I" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed ABSENCE, DUPLICATE ASSISTANCE, CHILD AGE"
				If mfip_elig_membs_code(elig_memb_count) = "J" Then mfip_elig_membs_status_info(elig_memb_count) = "Ineligible - Failed MFIP PERSON POST 60 REMOVAL"
				If mfip_elig_membs_code(elig_memb_count) = "N" Then mfip_elig_membs_status_info(elig_memb_count) = "Not a Unit Member"
				If mfip_elig_membs_code(elig_memb_count) = "A" Then mfip_elig_membs_deemed(elig_memb_count) = "Unit Member"
				If mfip_elig_membs_code(elig_memb_count) = "F" or mfip_elig_membs_code(elig_memb_count) = "G" or mfip_elig_membs_code(elig_memb_count) = "H" or mfip_elig_membs_code(elig_memb_count) = "J" Then mfip_elig_membs_deemed(elig_memb_count) = "Deemed"
				If mfip_elig_membs_code(elig_memb_count) = "D" or mfip_elig_membs_code(elig_memb_count) = "I" or mfip_elig_membs_code(elig_memb_count) = "N" Then mfip_elig_membs_deemed(elig_memb_count) = "Not Deemed"
				mfip_elig_membs_counted(elig_memb_count) = trim(mfip_elig_membs_counted(elig_memb_count))
				mfip_elig_membs_eligibility(elig_memb_count) = trim(mfip_elig_membs_eligibility(elig_memb_count))
				If mfip_elig_membs_budget_cycle(elig_memb_count) = "P" Then mfip_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
				If mfip_elig_membs_budget_cycle(elig_memb_count) = "R" Then mfip_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

				Call write_value_and_transmit("X", row, 3)
				EMReadScreen mfip_elig_membs_absence(elig_memb_count), 			6, 7, 17
				EMReadScreen mfip_elig_membs_child_age(elig_memb_count), 		6, 8, 17
				EMReadScreen mfip_elig_membs_citizenship(elig_memb_count), 		6, 9, 17
				EMReadScreen mfip_elig_membs_citizenship_verif(elig_memb_count),6, 10, 17
				EMReadScreen mfip_elig_membs_dupl_assist(elig_memb_count), 		6, 11, 17
				EMReadScreen mfip_elig_membs_foster_care(elig_memb_count), 		6, 12, 17
				EMReadScreen mfip_elig_membs_fraud(elig_memb_count), 			6, 13, 17
				EMReadScreen mfip_elig_membs_fs_disq(elig_memb_count), 			6, 17, 17

				mfip_elig_membs_absence(elig_memb_count) = trim(mfip_elig_membs_absence(elig_memb_count))
				mfip_elig_membs_child_age(elig_memb_count) = trim(mfip_elig_membs_child_age(elig_memb_count))
				mfip_elig_membs_citizenship(elig_memb_count) = trim(mfip_elig_membs_citizenship(elig_memb_count))
				mfip_elig_membs_citizenship_verif(elig_memb_count) = trim(mfip_elig_membs_citizenship_verif(elig_memb_count))
				mfip_elig_membs_dupl_assist(elig_memb_count) = trim(mfip_elig_membs_dupl_assist(elig_memb_count))
				mfip_elig_membs_foster_care(elig_memb_count) = trim(mfip_elig_membs_foster_care(elig_memb_count))
				mfip_elig_membs_fraud(elig_memb_count) = trim(mfip_elig_membs_fraud(elig_memb_count))
				mfip_elig_membs_fs_disq(elig_memb_count) = trim(mfip_elig_membs_fs_disq(elig_memb_count))


				EMReadScreen mfip_elig_membs_minor_living_arngmt(elig_memb_count), 	6, 7, 52
				EMReadScreen mfip_elig_membs_post_60_removal(elig_memb_count), 		6, 8, 52
				EMReadScreen mfip_elig_membs_ssi(elig_memb_count), 					6, 9, 52
				EMReadScreen mfip_elig_membs_ssn_code(elig_memb_count), 			6, 10, 52
				EMReadScreen mfip_elig_membs_unit_memb(elig_memb_count), 			6, 11, 52
				EMReadScreen mfip_elig_membs_unlawful_conduct(elig_memb_count), 	6, 12, 52
				EMReadScreen mfip_elig_membs_fs_recvd(elig_memb_count), 			6, 17, 52

				mfip_elig_membs_minor_living_arngmt(elig_memb_count) = trim(mfip_elig_membs_minor_living_arngmt(elig_memb_count))
				mfip_elig_membs_post_60_removal(elig_memb_count) = trim(mfip_elig_membs_post_60_removal(elig_memb_count))
				mfip_elig_membs_ssi(elig_memb_count) = trim(mfip_elig_membs_ssi(elig_memb_count))
				mfip_elig_membs_ssn_code(elig_memb_count) = trim(mfip_elig_membs_ssn_code(elig_memb_count))
				mfip_elig_membs_unit_memb(elig_memb_count) = trim(mfip_elig_membs_unit_memb(elig_memb_count))
				mfip_elig_membs_unlawful_conduct(elig_memb_count) = trim(mfip_elig_membs_unlawful_conduct(elig_memb_count))
				mfip_elig_membs_fs_recvd(elig_memb_count) = trim(mfip_elig_membs_fs_recvd(elig_memb_count))

				transmit

				Call write_value_and_transmit("X", row, 64)
				EMReadScreen emps_exists_for_memb, 19, 24, 2
				If emps_exists_for_memb = "EMPS DOES NOT EXIST" Then
					EMWriteScreen " ", row, 64
				Else
					EMReadScreen mfip_elig_membs_es_status_code(elig_memb_count), 2, 9, 22
					EMReadScreen mfip_elig_membs_es_status_info(elig_memb_count), 30, 9, 25
					mfip_elig_membs_es_status_code(elig_memb_count) = trim(mfip_elig_membs_es_status_code(elig_memb_count))
					mfip_elig_membs_es_status_info(elig_memb_count) = trim(mfip_elig_membs_es_status_info(elig_memb_count))
					transmit
				End If


				row = row + 1
				EMReadScreen next_ref_numb, 2, row, 6
				' MsgBox "row: " & row
			Loop until next_ref_numb = "  "

			transmit			'MFCR
			' MsgBox "In MFCR"

			EMReadScreen mfip_case_test_appl_withdraw, 		6, 6, 7
			EMReadScreen mfip_case_test_asset, 				6, 7, 7
			EMReadScreen mfip_case_test_death_applicant, 	6, 8, 7
			EMReadScreen mfip_case_test_dupl_assist, 		6, 9, 7
			EMReadScreen mfip_case_test_elig_child, 		6, 10, 7
			EMReadScreen mfip_case_test_fail_coop, 			6, 11, 7
			EMReadScreen mfip_case_test_fail_file, 			6, 12, 7
			EMReadScreen mfip_case_test_initial_income, 	6, 13, 7
			EMReadScreen mfip_case_test_minor_liv_arrange, 	6, 14, 7

			EMReadScreen mfip_case_test_monthly_income, 		6, 6, 46
			EMReadScreen mfip_case_test_post_60_disq, 			6, 7, 46
			EMReadScreen mfip_case_test_residence, 				6, 8, 46
			EMReadScreen mfip_case_test_sanction_limit, 		6, 9, 46
			EMReadScreen mfip_case_test_strike, 				6, 10, 46
			EMReadScreen mfip_case_test_TANF_time_limit, 		6, 11, 46
			EMReadScreen mfip_case_test_transfer_asset, 		6, 12, 46
			EMReadScreen mfip_case_test_verif, 					6, 13, 46
			EMReadScreen mfip_case_test_275_new_spouse_income, 	6, 14, 46

			EMReadScreen mfip_fs_case_test_fail_coop_snap_qc, 		6, 17, 7
			EMReadScreen mfip_fs_case_test_opt_out_cash, 			6, 17, 46
			EMReadScreen mfip_fs_case_test_opt_out_housing_grant, 	6, 18, 46

			If mfip_fs_case_test_opt_out_cash = "FAILED" Then mfip_cash_opt_out = True
			If mfip_fs_case_test_opt_out_housing_grant = "FAILED" Then mfip_HG_opt_out = True

			mfip_case_test_appl_withdraw = trim(mfip_case_test_appl_withdraw)
			mfip_case_test_asset = trim(mfip_case_test_asset)
			mfip_case_test_death_applicant = trim(mfip_case_test_death_applicant)
			mfip_case_test_dupl_assist = trim(mfip_case_test_dupl_assist)
			mfip_case_test_elig_child = trim(mfip_case_test_elig_child)
			mfip_case_test_fail_coop = trim(mfip_case_test_fail_coop)
			mfip_case_test_fail_file = trim(mfip_case_test_fail_file)
			mfip_case_test_initial_income = trim(mfip_case_test_initial_income)
			mfip_case_test_minor_liv_arrange = trim(mfip_case_test_minor_liv_arrange)
			mfip_case_test_monthly_income = trim(mfip_case_test_monthly_income)
			mfip_case_test_post_60_disq = trim(mfip_case_test_post_60_disq)
			mfip_case_test_residence = trim(mfip_case_test_residence)
			mfip_case_test_sanction_limit = trim(mfip_case_test_sanction_limit)
			mfip_case_test_strike = trim(mfip_case_test_strike)
			mfip_case_test_TANF_time_limit = trim(mfip_case_test_TANF_time_limit)
			mfip_case_test_transfer_asset = trim(mfip_case_test_transfer_asset)
			mfip_case_test_verif = trim(mfip_case_test_verif)
			mfip_case_test_275_new_spouse_income = trim(mfip_case_test_275_new_spouse_income)
			mfip_fs_case_test_fail_coop_snap_qc = trim(mfip_fs_case_test_fail_coop_snap_qc)
			mfip_fs_case_test_opt_out_cash = trim(mfip_fs_case_test_opt_out_cash)
			mfip_fs_case_test_opt_out_housing_grant = trim(mfip_fs_case_test_opt_out_housing_grant)

			Call write_value_and_transmit("X", 7, 5)						'ASSETS
			' MsgBox "In Asset Pop-UP"
			EMReadScreen mfip_counted_asset_CASH, 	10, 6, 47
			EMReadScreen mfip_counted_asset_ACCT, 	10, 7, 47
			EMReadScreen mfip_counted_asset_SECU, 	10, 8, 47
			EMReadScreen mfip_counted_asset_CARS, 	10, 9, 47
			EMReadScreen mfip_counted_asset_SPON, 	10, 10, 47
			EMReadScreen mfip_counted_asset_total, 	10, 12, 47
			EMReadScreen mfip_counted_asset_max, 	10, 13, 47

			mfip_counted_asset_CASH = trim(mfip_counted_asset_CASH)
			mfip_counted_asset_ACCT = trim(mfip_counted_asset_ACCT)
			mfip_counted_asset_SECU = trim(mfip_counted_asset_SECU)
			mfip_counted_asset_CARS = trim(mfip_counted_asset_CARS)
			mfip_counted_asset_SPON = trim(mfip_counted_asset_SPON)
			mfip_counted_asset_total = trim(mfip_counted_asset_total)
			mfip_counted_asset_max = trim(mfip_counted_asset_max)

			transmit
			' MsgBox "Back to MFCR - 1"

			Call write_value_and_transmit("X", 13, 5)						'INITIAL INCOME
			' MsgBox "In Initial Income Pop_up"
			EMReadScreen mfip_initial_income_earned, 			10, 8, 51
			EMReadScreen mfip_initial_income_deoendant_care, 	10, 9, 51
			EMReadScreen mfip_initial_income_unearned, 			10, 10, 51
			EMReadScreen mfip_initial_income_deemed, 			10, 11, 51
			EMReadScreen mfip_initial_income_cses_exclusion, 	10, 12, 51
			EMReadScreen mfip_initial_income_total, 			10, 13, 51
			EMReadScreen mfip_initial_income_family_wage_level, 10, 15, 51

			mfip_initial_income_earned = trim(mfip_initial_income_earned)
			mfip_initial_income_deoendant_care = trim(mfip_initial_income_deoendant_care)
			mfip_initial_income_unearned = trim(mfip_initial_income_unearned)
			mfip_initial_income_deemed = trim(mfip_initial_income_deemed)
			mfip_initial_income_cses_exclusion = trim(mfip_initial_income_cses_exclusion)
			mfip_initial_income_total = trim(mfip_initial_income_total)
			mfip_initial_income_family_wage_level = trim(mfip_initial_income_family_wage_level)

			'TODO - Read each person's information in the pop-ups
			Call write_value_and_transmit("X", 8, 20)		'Member Initial Earned Income
			' MsgBox "Member Initial Earned Income"
			Do
				EMReadScreen pop_up_name, 40, 8, 28
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then
						EMReadScreen mfip_elig_membs_initial_BUSI_inc_total(case_memb), 		10, 11, 54
						EMReadScreen mfip_elig_membs_initial_JOBS_inc_total(case_memb), 		10, 12, 54
						EMReadScreen mfip_elig_membs_initial_earned_inc_total(case_memb), 		10, 13, 54
						EMReadScreen mfip_elig_membs_initial_stndrd_ei_disregard(case_memb), 	10, 14, 54
						EMReadScreen mfip_elig_membs_initial_earned_inc_subtotal(case_memb), 	10, 15, 54
						EMReadScreen mfip_elig_membs_initial_earned_inc_disregard(case_memb), 	10, 16, 54
						EMReadScreen mfip_elig_membs_initial_avail_earned_inc(case_memb), 		10, 17, 54
						EMReadScreen mfip_elig_membs_initial_allocation(case_memb), 			10, 18, 54
						EMReadScreen mfip_elig_membs_initial_child_support(case_memb), 			10, 19, 54
						EMReadScreen mfip_elig_membs_initial_counted_earned_inc_total(case_memb), 10, 20, 54

						mfip_elig_membs_initial_BUSI_inc_total(case_memb) = trim(mfip_elig_membs_initial_BUSI_inc_total(case_memb))
						mfip_elig_membs_initial_JOBS_inc_total(case_memb) = trim(mfip_elig_membs_initial_JOBS_inc_total(case_memb))
						mfip_elig_membs_initial_earned_inc_total(case_memb) = trim(mfip_elig_membs_initial_earned_inc_total(case_memb))
						mfip_elig_membs_initial_stndrd_ei_disregard(case_memb) = trim(mfip_elig_membs_initial_stndrd_ei_disregard(case_memb))
						mfip_elig_membs_initial_earned_inc_subtotal(case_memb) = trim(mfip_elig_membs_initial_earned_inc_subtotal(case_memb))
						mfip_elig_membs_initial_earned_inc_disregard(case_memb) = trim(mfip_elig_membs_initial_earned_inc_disregard(case_memb))
						mfip_elig_membs_initial_avail_earned_inc(case_memb) = trim(mfip_elig_membs_initial_avail_earned_inc(case_memb))
						mfip_elig_membs_initial_allocation(case_memb) = trim(mfip_elig_membs_initial_allocation(case_memb))
						mfip_elig_membs_initial_child_support(case_memb) = trim(mfip_elig_membs_initial_child_support(case_memb))
						mfip_elig_membs_initial_counted_earned_inc_total(case_memb) = trim(mfip_elig_membs_initial_counted_earned_inc_total(case_memb))

						' If mfip_elig_membs_initial_BUSI_inc_total(case_memb) <> "0.00" Then 			'this will likely not be used - opening these pop ups do not provide details on different jobs
						' 	Call write_value_and_transmit("X", 11, 20)
						' End If
						' If mfip_elig_membs_initial_JOBS_inc_total(case_memb) <> "0.00" Then
						' 	Call write_value_and_transmit("X", 12, 20)
						' End If
					End If
				Next
				transmit

				EMReadScreen back_to_menu, 14, 6, 29
			Loop until back_to_menu = "Initial Income"
			' MsgBox "Back to the Initial Income Pop-Up"

			If mfip_initial_income_deoendant_care <> "0.00" Then 			''Depended Care Initial Income calculation pop-up
				Call write_value_and_transmit("X", 9, 20)
			End If

			Call write_value_and_transmit("X", 10, 20)		'Member Initial Unearned Income
			' MsgBox "Member Initial Unearned Income"
			Do
				EMReadScreen pop_up_name, 40, 8, 28
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then
						EMReadScreen mfip_elig_membs_initial_UNEA_inc_total(case_memb), 		10, 11, 49
						EMReadScreen mfip_elig_membs_initial_allocation_balance(case_memb), 	10, 12, 49
						EMReadScreen mfip_elig_membs_initial_child_support_balance(case_memb), 	10, 13, 49
						EMReadScreen mfip_elig_membs_initial_counted_UNEA_inc_total(case_memb), 10, 14, 49

						mfip_elig_membs_initial_UNEA_inc_total(case_memb) = trim(mfip_elig_membs_initial_UNEA_inc_total(case_memb))
						mfip_elig_membs_initial_allocation_balance(case_memb) = trim(mfip_elig_membs_initial_allocation_balance(case_memb))
						mfip_elig_membs_initial_child_support_balance(case_memb) = trim(mfip_elig_membs_initial_child_support_balance(case_memb))
						 mfip_elig_membs_initial_counted_UNEA_inc_total(case_memb) = trim(mfip_elig_membs_initial_counted_UNEA_inc_total(case_memb))
					End If
				Next
				transmit

				EMReadScreen back_to_menu, 14, 6, 29
			Loop until back_to_menu = "Initial Income"
			' MsgBox "Back to the Initial Income Pop-Up"

			If mfip_initial_income_deemed <> "0.00" Then 			'Deemed Initial Income calculation pop-up
				Call write_value_and_transmit("X", 11, 20)

				Do
					EMReadScreen pop_up_name, 40, 8, 28
					pop_up_name = trim(pop_up_name)
					For case_memb = 0 to UBound(mfip_elig_ref_numbs)
						If pop_up_name = mfip_elig_membs_full_name(case_memb) Then
							' EMReadScreen mfip_elig_membs_deemer_initial_BUSI_inc_total(case_memb), 			9, 9, 52
							' EMReadScreen mfip_elig_membs_deemer_initial_JOBS_inc_total(case_memb), 			9, 10, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_earned_inc_total(case_memb), 		9, 11, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_stndrd_ei_disregard(case_memb), 	9, 12, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_earned_inc_subtotal(case_memb), 	9, 13, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_earned_inc_disregard(case_memb), 	9, 14, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_unearned_inc(case_memb), 			9, 15, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_sub_total_counted_income(case_memb),9, 17, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_allocation(case_memb), 				9, 18, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_child_support(case_memb), 			9, 19, 54
							' EMReadScreen mfip_elig_membs_deemer_initial_counted_income_total(case_memb),	9, 20, 54
							'
							' mfip_elig_membs_deemer_initial_BUSI_inc_total(case_memb) = trim(mfip_elig_membs_deemer_initial_BUSI_inc_total(case_memb))
							' mfip_elig_membs_deemer_initial_JOBS_inc_total(case_memb) = trim(mfip_elig_membs_deemer_initial_JOBS_inc_total(case_memb))
							' mfip_elig_membs_deemer_initial_earned_inc_total(case_memb) = trim(mfip_elig_membs_deemer_initial_earned_inc_total(case_memb))
							' mfip_elig_membs_deemer_initial_stndrd_ei_disregard(case_memb) = trim(mfip_elig_membs_deemer_initial_stndrd_ei_disregard(case_memb))
							' mfip_elig_membs_deemer_initial_earned_inc_subtotal(case_memb) = trim(mfip_elig_membs_deemer_initial_earned_inc_subtotal(case_memb))
							' mfip_elig_membs_deemer_initial_earned_inc_disregard(case_memb) = trim(mfip_elig_membs_deemer_initial_earned_inc_disregard(case_memb))
							' mfip_elig_membs_deemer_initial_unearned_inc(case_memb) = trim(mfip_elig_membs_deemer_initial_unearned_inc(case_memb))
							' mfip_elig_membs_deemer_initial_sub_total_counted_income(case_memb) = trim(mfip_elig_membs_deemer_initial_sub_total_counted_income(case_memb))
							' mfip_elig_membs_deemer_initial_allocation(case_memb) = trim(mfip_elig_membs_deemer_initial_allocation(case_memb))
							' mfip_elig_membs_deemer_initial_child_support(case_memb) = trim(mfip_elig_membs_deemer_initial_child_support(case_memb))
							' mfip_elig_membs_deemer_initial_counted_income_total(case_memb) = trim(mfip_elig_membs_deemer_initial_counted_income_total(case_memb))

						End If
					Next
					transmit

					EMReadScreen back_to_menu, 14, 6, 29
				Loop until back_to_menu = "Initial Income"
			End If

			Call write_value_and_transmit("X", 12, 20)				'CSES Exclusion Initiall Income calculation pop-up
			EMWaitReady 0, 0
			' MsgBox "CSES Exclusion Pop-Up"
			EMReadScreen mfip_initial_income_cses_income, 10, 9, 52
			EMReadScreen mfip_initial_income_cses_child_count, 2, 11, 37

			mfip_initial_income_cses_income = trim(mfip_initial_income_cses_income)
			mfip_initial_income_cses_child_count = trim(mfip_initial_income_cses_child_count)

			Call write_value_and_transmit("X", 9, 20)				'open cses initial income pop-up'
			' MsgBox "CSES initial Income"

			EMReadScreen mfip_initial_income_net_cses_income, 10, 19, 44
			mfip_initial_income_net_cses_income = trim(mfip_initial_income_net_cses_income)
			mfcr_row = 7
			Do
				EMReadScreen ref_numb, 2, mfcr_row, 7

				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If ref_numb = mfip_elig_ref_numbs(case_memb) Then
						EMReadScreen mfip_elig_membs_initial_income_cses_retro_income(case_memb), 10, mfcr_row, 41
						EMReadScreen mfip_elig_membs_initial_income_cses_prosp_income(case_memb), 10, mfcr_row, 54

						mfip_elig_membs_initial_income_cses_retro_income(case_memb) = trim(mfip_elig_membs_initial_income_cses_retro_income(case_memb))
						mfip_elig_membs_initial_income_cses_prosp_income(case_memb) = trim(mfip_elig_membs_initial_income_cses_prosp_income(case_memb))
					End If
				Next

				mfcr_row = mfcr_row + 1
				EMReadScreen next_ref_numb, 2, mfcr_row, 3
			Loop until next_ref_numb = "  "

			PF3			'back to CSES Exclusion caclulaiton
			' MsgBox "back to CSES Exclusion calculation"
			PF3			'back to initial income calculation
			' MsgBox "Back to Initial Income Pop-Up"
			PF3			'back to main mf elig panel'
			' MsgBox "Back to MFCR - 2"

			Call write_value_and_transmit("X", 14, 44)						'NEW SPOUSE 275% INCOME
			EMReadScreen mfip_12_month_start_date, 				8, 6, 46
			EMReadScreen mfip_designated_spouse_ref_numb, 		2, 7, 46
			EMReadScreen mfip_new_spouse_inc_earned, 			10, 11, 57
			EMReadScreen mfip_new_spouse_inc_unearned, 			10, 12, 57
			EMReadScreen mfip_new_spouse_inc_deemed_earned, 	10, 13, 57
			EMReadScreen mfip_new_spouse_inc_deemed_unearned, 	10, 14, 57
			EMReadScreen mfip_new_spouse_inc_total, 			10, 16, 57
			EMReadScreen mfip_275_fpg_amt, 						10, 18, 57
			EMReadScreen mfip_hh_size_fornew_spouse_calc, 		2, 18, 51

			mfip_12_month_start_date = trim(mfip_12_month_start_date)
			mfip_designated_spouse_ref_numb = trim(mfip_designated_spouse_ref_numb)
			mfip_new_spouse_inc_earned = trim(mfip_new_spouse_inc_earned)
			mfip_new_spouse_inc_unearned = trim(mfip_new_spouse_inc_unearned)
			mfip_new_spouse_inc_deemed_earned = trim(mfip_new_spouse_inc_deemed_earned)
			mfip_new_spouse_inc_deemed_unearned = trim(mfip_new_spouse_inc_deemed_unearned)
			mfip_new_spouse_inc_total = trim(mfip_new_spouse_inc_total)
			mfip_275_fpg_amt = trim(mfip_275_fpg_amt)
			mfip_hh_size_fornew_spouse_calc = trim(mfip_hh_size_fornew_spouse_calc)

			Call write_value_and_transmit("X", 11, 20)		'Member earned and unearned for New Spouse calculation
			Do
				EMReadScreen pop_up_name, 35, 7, 25
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

						EMReadScreen mfip_elig_membs_new_spouse_earned_income(case_memb), 	10, 9, 48
						EMReadScreen mfip_elig_membs_new_spouse_unearned_income(case_memb), 10, 10, 48
						EMReadScreen mfip_elig_membs_new_spouse_total_income(case_memb), 	10, 11, 48

						mfip_elig_membs_new_spouse_earned_income(case_memb) = trim(mfip_elig_membs_new_spouse_earned_income(case_memb))
						mfip_elig_membs_new_spouse_unearned_income(case_memb) = trim(mfip_elig_membs_new_spouse_unearned_income(case_memb))
						mfip_elig_membs_new_spouse_total_income(case_memb) = trim(mfip_elig_membs_new_spouse_total_income(case_memb))
					End If
				Next
				transmit

				EMReadScreen back_to_menu, 17, 7, 22
			Loop until back_to_menu = "Designated Spouse"

			'TODO - Read the deemed pop-ups
			If mfip_new_spouse_inc_deemed_earned <> "0.00" Then
				' Call write_value_and_transmit("X", 13, 20)		'Member deemed earned for New Spouse calculation
			End If
			If mfip_new_spouse_inc_deemed_unearned <> "0.00" Then
				' Call write_value_and_transmit("X", 14, 20)		'Member deemed unearned for New Spouse calculation
			End If

			PF3
			' MsgBox "Back to MFCR - 3"


			transmit			'MFBF
			' MsgBox "In MFBF"
			mfbf_row = 7
			Do
				EMReadScreen ref_numb, 2, mfbf_row, 3

				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If ref_numb = mfip_elig_ref_numbs(case_memb) Then
						EMReadScreen mfip_memb_cash_portion_code(case_memb), 	1, mfbf_row, 37
						EMReadScreen mfip_memb_food_portion_code(case_memb), 	1, mfbf_row, 45
						EMReadScreen mfip_memb_state_food_code(case_memb), 		1, mfbf_row, 54
						EMReadScreen mfip_memb_sanction_yn(case_memb), 			1, mfbf_row, 68
						If mfip_memb_sanction_yn(case_memb) = "Y" Then mfip_case_in_sancttion = True

						Call write_value_and_transmit("X", mfbf_row, 62)
						EMReadScreen mfip_memb_sanction_child_support_test(case_memb),	6, 7, 12
						EMReadScreen mfip_memb_sanction_drug_felon_test(case_memb), 	6, 7, 12
						EMReadScreen mfip_memb_sanction_emp_services_test(case_memb), 	6, 7, 12
						EMReadScreen mfip_memb_sanction_fin_orient_test(case_memb), 	6, 7, 12

						EMReadScreen mfip_memb_sanction_occurence(case_memb), 1, 12, 21
						EMReadScreen mfip_memb_sanction_begin_date(case_memb), 7, 12, 40
						EMReadScreen mfip_memb_sanction_last_sanc_month(case_memb), 55, 12, 62
						transmit
					End If
				Next

				mfbf_row = mfbf_row + 1
				EMReadScreen next_ref_numb, 2, mfbf_row, 3
			Loop until next_ref_numb = "  "

			EMReadScreen mfip_case_sanction_percent, 3, 18, 28
			EMReadScreen mfip_case_sanction_vendor_yn, 1, 18, 48
			EMReadScreen mfip_case_sanction_last_vendor_month, 7, 18, 68

			mfip_case_sanction_percent = trim(mfip_case_sanction_percent)
			mfip_case_sanction_vendor_yn = trim(mfip_case_sanction_vendor_yn)
			mfip_case_sanction_last_vendor_month = trim(mfip_case_sanction_last_vendor_month)

			transmit			'MFB1
			' MsgBox "In MFB1"
			EMReadScreen mfip_case_budg_family_wage_level, 				10, 5, 32
			EMReadScreen mfip_case_budg_monthly_earned_income, 			10, 6, 32
			EMReadScreen mfip_case_budg_wage_level_earned_inc_difference, 10, 7, 32
			EMReadScreen mfip_case_budg_transitional_standard, 			10, 9, 32
			EMReadScreen mfip_case_budg_monthly_need, 					10, 10, 32
			EMReadScreen mfip_case_budg_unearned_income, 				10, 11, 32
			EMReadScreen mfip_case_budg_deemed_income, 					10, 12, 32
			EMReadScreen mfip_case_budg_cses_exclusion, 				10, 13, 32
			EMReadScreen mfip_case_budg_unmet_need, 					10, 14, 32
			EMReadScreen mfip_case_budg_unmet_need_food_potion, 		10, 15, 32
			EMReadScreen mfip_case_budg_tribal_counted_income, 			10, 16, 32
			EMReadScreen mfip_case_budg_unmet_need_cash_portion, 		10, 17, 32
			EMReadScreen mfip_case_budg_deduction_subsidy_tribal_cses, 	10, 18, 32


			EMReadScreen mfip_case_budg_net_food_portion, 			10, 5, 71
			EMReadScreen mfip_case_budg_net_cash_portion, 			10, 6, 71
			EMReadScreen mfip_case_budg_net_unmet_need, 			10, 7, 71
			EMReadScreen mfip_case_budg_deduction_sanction_vendor, 	10, 8, 71
			EMReadScreen mfip_case_budg_unmet_neet_subtotal, 		10, 9, 71
			EMReadScreen mfip_case_budg_subtotal_food_portion, 		10, 11, 71
			EMReadScreen mfip_case_budg_food_portion_deduction, 	10, 12, 71
			EMReadScreen mfip_case_budg_entitlement_food_portion, 	10, 13, 71
			EMReadScreen mfip_case_budg_entitlement_housing_grant, 	10, 15, 71

			mfip_case_budg_family_wage_level = trim(mfip_case_budg_family_wage_level)
			mfip_case_budg_monthly_earned_income = trim(mfip_case_budg_monthly_earned_income)
			mfip_case_budg_wage_level_earned_inc_difference = trim(mfip_case_budg_wage_level_earned_inc_difference)
			mfip_case_budg_transitional_standard = trim(mfip_case_budg_transitional_standard)
			mfip_case_budg_monthly_need = trim(mfip_case_budg_monthly_need)
			mfip_case_budg_unearned_income = trim(mfip_case_budg_unearned_income)
			mfip_case_budg_deemed_income = trim(mfip_case_budg_deemed_income)
			mfip_case_budg_cses_exclusion = trim(mfip_case_budg_cses_exclusion)
			mfip_case_budg_unmet_need = trim(mfip_case_budg_unmet_need)
			mfip_case_budg_unmet_need_food_potion = trim(mfip_case_budg_unmet_need_food_potion)
			mfip_case_budg_tribal_counted_income = trim(mfip_case_budg_tribal_counted_income)
			mfip_case_budg_unmet_need_cash_portion = trim(mfip_case_budg_unmet_need_cash_portion)
			mfip_case_budg_deduction_subsidy_tribal_cses = trim(mfip_case_budg_deduction_subsidy_tribal_cses)

			mfip_case_budg_net_food_portion = trim(mfip_case_budg_net_food_portion)
			mfip_case_budg_net_cash_portion = trim(mfip_case_budg_net_cash_portion)
			mfip_case_budg_net_unmet_need = trim(mfip_case_budg_net_unmet_need)
			mfip_case_budg_deduction_sanction_vendor = trim(mfip_case_budg_deduction_sanction_vendor)
			mfip_case_budg_unmet_neet_subtotal = trim(mfip_case_budg_unmet_neet_subtotal)
			mfip_case_budg_subtotal_food_portion = trim(mfip_case_budg_subtotal_food_portion)
			mfip_case_budg_food_portion_deduction = trim(mfip_case_budg_food_portion_deduction)
			mfip_case_budg_entitlement_food_portion = trim(mfip_case_budg_entitlement_food_portion)
			mfip_case_budg_entitlement_housing_grant = trim(mfip_case_budg_entitlement_housing_grant)

			Call write_value_and_transmit("X", 6, 3)		' member specific EARNED INCOME
			Do
				EMReadScreen pop_up_name, 40, 8, 28
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

						EMReadScreen mfip_elig_membs_self_emp_income(case_memb), 				10, 11, 54
						EMReadScreen mfip_elig_membs_wages_income(case_memb), 					10, 12, 54
						EMReadScreen mfip_elig_membs_total_earned_income(case_memb), 			10, 13, 54
						EMReadScreen mfip_elig_membs_standard_EI_disregard(case_memb), 			10, 14, 54
						EMReadScreen mfip_elig_membs_earned_income_subtotal(case_memb), 		10, 15, 54
						EMReadScreen mfip_elig_membs_earned_income_50_perc_disregard(case_memb), 10, 16, 54
						EMReadScreen mfip_elig_membs_available_earned_income(case_memb), 		10, 17, 54
						EMReadScreen mfip_elig_membs_allocation_deduction(case_memb), 			10, 18, 54
						EMReadScreen mfip_elig_membs_child_support_deduction(case_memb), 		10, 19, 54
						EMReadScreen mfip_elig_membs_counted_earned_income(case_memb), 			10, 20, 54

						mfip_elig_membs_self_emp_income(case_memb) = trim(mfip_elig_membs_self_emp_income(case_memb))
						mfip_elig_membs_wages_income(case_memb) = trim(mfip_elig_membs_wages_income(case_memb))
						mfip_elig_membs_total_earned_income(case_memb) = trim(mfip_elig_membs_total_earned_income(case_memb))
						mfip_elig_membs_standard_EI_disregard(case_memb) = trim(mfip_elig_membs_standard_EI_disregard(case_memb))
						mfip_elig_membs_earned_income_subtotal(case_memb) = trim(mfip_elig_membs_earned_income_subtotal(case_memb))
						mfip_elig_membs_earned_income_50_perc_disregard(case_memb) = trim(mfip_elig_membs_earned_income_50_perc_disregard(case_memb))
						mfip_elig_membs_available_earned_income(case_memb) = trim(mfip_elig_membs_available_earned_income(case_memb))
						mfip_elig_membs_allocation_deduction(case_memb) = trim(mfip_elig_membs_allocation_deduction(case_memb))
						mfip_elig_membs_child_support_deduction(case_memb) = trim(mfip_elig_membs_child_support_deduction(case_memb))
						mfip_elig_membs_counted_earned_income(case_memb) = trim(mfip_elig_membs_counted_earned_income(case_memb))

					End If
				Next
				transmit
				EMReadScreen still_in_menu, 12, 5, 32
			Loop until still_in_menu <> "Maxis Person"

			Call write_value_and_transmit("X", 11, 3)		' member specific UNEARNED INCOME
			Do
				EMReadScreen pop_up_name, 25, 8, 34
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

						EMReadScreen mfip_elig_membs_total_unearned_income(case_memb), 	10, 11, 54
						EMReadScreen mfip_elig_membs_allocation_balance(case_memb), 	10, 12, 54
						EMReadScreen mfip_elig_membs_child_support_balance(case_memb), 	10, 13, 54
						EMReadScreen mfip_elig_membs_counted_unearned_income(case_memb), 10, 14, 54

						mfip_elig_membs_total_unearned_income(case_memb) = trim(mfip_elig_membs_total_unearned_income(case_memb))
						mfip_elig_membs_allocation_balance(case_memb) = trim(mfip_elig_membs_allocation_balance(case_memb))
						mfip_elig_membs_child_support_balance(case_memb) = trim(mfip_elig_membs_child_support_balance(case_memb))
						mfip_elig_membs_counted_unearned_income(case_memb) = trim(mfip_elig_membs_counted_unearned_income(case_memb))

					End If
				Next
				transmit
				EMReadScreen still_in_menu, 15, 6, 34
			Loop until still_in_menu <> "Unearned Income"

			' Call write_value_and_transmit("X", 12, 3)		'TODO member specific DEEMED INCOME

			Call write_value_and_transmit("X", 13, 3)		'Child Support Exclusion'
			EMReadScreen mfip_budg_cses_excln_cses_income, 10, 9, 52
			EMReadScreen mfip_budg_cses_excln_child_count, 2, 11, 37
			EMReadScreen mfip_budg_cses_excln_total, 10, 13, 52

			mfip_budg_cses_excln_cses_income = trim(mfip_budg_cses_excln_cses_income)
			mfip_budg_cses_excln_child_count = trim(mfip_budg_cses_excln_child_count)
			mfip_budg_cses_excln_total = trim(mfip_budg_cses_excln_total)

			transmit

			Call write_value_and_transmit("X", 16, 5)		' member specific TRIBAL INCOME
			EMReadScreen mfip_budg_total_county_88_child_support_income, 	10, 6, 55
			EMReadScreen mfip_budg_total_county_88_gaming_income, 			10, 7, 55
			EMReadScreen mfip_budg_total_tribal_income_fs_portion_deduction, 10, 8, 55
			mfip_budg_total_county_88_child_support_income = trim(mfip_budg_total_county_88_child_support_income)
			mfip_budg_total_county_88_gaming__income = trim(mfip_budg_total_county_88_gaming__income)
			mfip_budg_total_tribal_income_fs_portion_deduction = trim(mfip_budg_total_tribal_income_fs_portion_deduction)

			Call write_value_and_transmit("X", 6, 12)		' member specific Tribal Child Support Income
			Do
				EMReadScreen pop_up_name, 25, 8, 34
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

						EMReadScreen mfip_elig_membs_county_88_cses_income(case_memb), 10, 11, 54

						mfip_elig_membs_county_88_cses_income(case_memb) = trim(mfip_elig_membs_county_88_cses_income(case_memb))
					End If
				Next
				transmit
				EMReadScreen back_to_menu, 21, 4, 31
			Loop until back_to_menu = "Tribal Counted Income"

			Call write_value_and_transmit("X", 7, 12)		' member specific Tribal Gaming Income
			Do
				EMReadScreen pop_up_name, 30, 7, 37
				pop_up_name = trim(pop_up_name)
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If pop_up_name = mfip_elig_membs_full_name(case_memb) Then

						EMReadScreen mfip_elig_membs_county_88_gaming_income(case_memb), 	10, 10, 61
						EMReadScreen mfip_elig_membs_county_88_200_perc_fpg(case_memb), 	10, 11, 61
						EMReadScreen mfip_elig_membs_county_88_deemers_unmet_need(case_memb), 10, 12, 61
						EMReadScreen mfip_elig_membs_county_88_allocation(case_memb), 		10, 13, 61
						EMReadScreen mfip_elig_membs_county_88_child_support(case_memb), 	10, 14, 61
						EMReadScreen mfip_elig_membs_county_88_counted_gaming_income(case_memb), 10, 15, 61

						mfip_elig_membs_county_88_gaming_income(case_memb) = trim(mfip_elig_membs_county_88_gaming_income(case_memb))
						mfip_elig_membs_county_88_200_perc_fpg(case_memb) = trim(mfip_elig_membs_county_88_200_perc_fpg(case_memb))
						mfip_elig_membs_county_88_deemers_unmet_need(case_memb) = trim(mfip_elig_membs_county_88_deemers_unmet_need(case_memb))
						mfip_elig_membs_county_88_allocation(case_memb) = trim(mfip_elig_membs_county_88_allocation(case_memb))
						mfip_elig_membs_county_88_child_support(case_memb) = trim(mfip_elig_membs_county_88_child_support(case_memb))
						mfip_elig_membs_county_88_counted_gaming_income(case_memb) = trim(mfip_elig_membs_county_88_counted_gaming_income(case_memb))
					End If
				Next
				transmit
				EMReadScreen back_to_menu, 21, 4, 31
			Loop until back_to_menu = "Tribal Counted Income"
			transmit                  ''back to MFB1

			Call write_value_and_transmit("X", 18, 5)		' member specific SUBSIDY
			EMReadScreen mfip_budg_total_housing_subsidy_amount, 10, 8, 51
			EMReadScreen mfip_budg_total_tribal_child_support, 10, 9, 51
			EMReadScreen mfip_budg_total_subsidy_tribal_cash_portion_deduction, 10, 10, 51
			mfip_budg_total_housing_subsidy_amount = trim(mfip_budg_total_housing_subsidy_amount)
			mfip_budg_total_tribal_child_support = trim(mfip_budg_total_tribal_child_support)
			mfip_budg_total_subsidy_tribal_cash_portion_deduction = trim(mfip_budg_total_subsidy_tribal_cash_portion_deduction)

			Call write_value_and_transmit("X", 8, 13)		' member specific subsidy Income
			EMReadScreen mfip_elig_budg_total_countable_housing_subsidy, 10, 19, 48
			EMReadScreen mfip_elig_budg_housing_subsidy_exempt, 1, 21, 47

			mfip_elig_budg_total_countable_housing_subsidy = trim(mfip_elig_budg_total_countable_housing_subsidy)
			mfip_elig_budg_housing_subsidy_exempt = trim(mfip_elig_budg_housing_subsidy_exempt)

			Do
				row = 8
				EMReadScreen memb_ref_numb, 2, row, 6
				For case_memb = 0 to UBound(mfip_elig_ref_numbs)
					If memb_ref_numb = mfip_elig_ref_numbs(case_memb) Then

						EMReadScreen mfip_elig_membs_retro_subsidy_amount(case_memb), 10, row, 38
						EMReadScreen mfip_elig_membs_prosp_subsidy_amount(case_memb), 10, row, 49

						mfip_elig_membs_retro_subsidy_amount(case_memb) = trim(mfip_elig_membs_retro_subsidy_amount(case_memb))
						mfip_elig_membs_prosp_subsidy_amount(case_memb) = trim(mfip_elig_membs_prosp_subsidy_amount(case_memb))
					End If
				Next
				row = row + 1
				EMReadScreen next_memb_ref_numb, 2, row, 6
			Loop until next_memb_ref_numb = "  "
			transmit 					'back to pop-up

			transmit                 	'back to MFB1

			Call write_value_and_transmit("X", 8, 44)		'Sanction and Vendor
			EMReadScreen mfip_case_budg_10_perc_sanc, 					10, 7, 55
			EMReadScreen mfip_case_budg_unmet_need_after_pre_vndr_sanc, 10, 8, 55
			EMReadScreen mfip_case_budg_sanc_calc_food_portion, 		10, 9, 55
			EMReadScreen mfip_case_budg_sanc_calc_cash_portion, 		10, 10, 55
			EMReadScreen mfip_case_budg_pot_mand_vndr_pymts, 			10, 11, 55
			EMReadScreen mfip_case_budg_30_perc_sanc, 					10, 12, 55

			mfip_case_budg_10_perc_sanc = trim(mfip_case_budg_10_perc_sanc)
			mfip_case_budg_unmet_need_after_pre_vndr_sanc = trim(mfip_case_budg_unmet_need_after_pre_vndr_sanc)
			mfip_case_budg_sanc_calc_food_portion = trim(mfip_case_budg_sanc_calc_food_portion)
			mfip_case_budg_sanc_calc_cash_portion = trim(mfip_case_budg_sanc_calc_cash_portion)
			mfip_case_budg_pot_mand_vndr_pymts = trim(mfip_case_budg_pot_mand_vndr_pymts)
			mfip_case_budg_30_perc_sanc = trim(mfip_case_budg_30_perc_sanc)
			transmit

			Call write_value_and_transmit("X", 12, 44)		'Food portion Deduction
			EMReadScreen mfip_case_budg_non_citzn_fs_inelig_pers_count, 1, 10, 17
			EMReadScreen mfip_case_budg_non_citzn_fs_inelig_amt, 		10, 10, 45
			EMReadScreen mfip_case_budg_other_fs_inelig_pers_count, 	1, 12, 17
			EMReadScreen mfip_case_budg_other_fs_inelig_amt, 			10, 12, 45

			mfip_case_budg_non_citzn_fs_inelig_pers_count = trim(mfip_case_budg_non_citzn_fs_inelig_pers_count)
			mfip_case_budg_non_citzn_fs_inelig_amt = trim(mfip_case_budg_non_citzn_fs_inelig_amt)
			mfip_case_budg_other_fs_inelig_pers_count = trim(mfip_case_budg_other_fs_inelig_pers_count)
			mfip_case_budg_other_fs_inelig_amt = trim(mfip_case_budg_other_fs_inelig_amt)
			transmit

			transmit			'MFB2
			' MsgBox "In MFB2"
			EMReadScreen mfip_case_budg_prorate_date, 8, 5, 19

			EMReadScreen mfip_case_budg_fed_food_benefit, 			10, 7, 32
			EMReadScreen mfip_case_budg_food_prorated_amt, 			10, 8, 32
			EMReadScreen mfip_case_budg_entitlement_cash_portion, 	10, 10, 32
			EMReadScreen mfip_case_budg_mand_sanc_vendor, 			10, 11, 32
			EMReadScreen mfip_case_budg_net_cash_after_sanc_portion, 10, 12, 32
			EMReadScreen mfip_case_budg_cash_prorated_amt, 			10, 13, 32
			EMReadScreen mfip_case_budg_state_food_benefit, 		10, 15, 32
			EMReadScreen mfip_case_budg_state_food_prorated_amt, 	10, 16, 32
			' EMReadScreen mfip_case_budg_entitlement_cash_portion, 10, 10, 32

			EMReadScreen mfip_case_budg_grant_amount, 				10, 5, 71
			EMReadScreen mfip_case_budg_amt_already_issued, 		10, 8, 71
			EMReadScreen mfip_case_budg_supplement_due, 			10, 9, 71
			EMReadScreen mfip_case_budg_overpayment, 				10, 10, 71
			EMReadScreen mfip_case_budg_adjusted_grant_amt, 		10, 11, 71
			EMReadScreen mfip_case_budg_recoupment, 				10, 12, 71
			EMReadScreen mfip_case_budg_total_food_issuance, 		10, 14, 71
			EMReadScreen mfip_case_budg_total_cash_issuance, 		10, 15, 71
			EMReadScreen mfip_case_budg_total_housing_grant_issuance, 10, 16, 71

			mfip_case_budg_prorate_date = trim(mfip_case_budg_prorate_date)
			mfip_case_budg_fed_food_benefit = trim(mfip_case_budg_fed_food_benefit)
			mfip_case_budg_food_prorated_amt = trim(mfip_case_budg_food_prorated_amt)
			mfip_case_budg_entitlement_cash_portion = trim(mfip_case_budg_entitlement_cash_portion)
			mfip_case_budg_mand_sanc_vendor = trim(mfip_case_budg_mand_sanc_vendor)
			mfip_case_budg_net_cash_after_sanc_portion = trim(mfip_case_budg_net_cash_after_sanc_portion)
			mfip_case_budg_cash_prorated_amt = trim(mfip_case_budg_cash_prorated_amt)
			mfip_case_budg_state_food_benefit = trim(mfip_case_budg_state_food_benefit)
			mfip_case_budg_state_food_prorated_amt = trim(mfip_case_budg_state_food_prorated_amt)
			mfip_case_budg_grant_amount = trim(mfip_case_budg_grant_amount)
			mfip_case_budg_amt_already_issued = trim(mfip_case_budg_amt_already_issued)
			mfip_case_budg_supplement_due = trim(mfip_case_budg_supplement_due)
			mfip_case_budg_overpayment = trim(mfip_case_budg_overpayment)
			mfip_case_budg_adjusted_grant_amt = trim(mfip_case_budg_adjusted_grant_amt)
			mfip_case_budg_recoupment = trim(mfip_case_budg_recoupment)
			mfip_case_budg_total_food_issuance = trim(mfip_case_budg_total_food_issuance)
			mfip_case_budg_total_cash_issuance = trim(mfip_case_budg_total_cash_issuance)
			mfip_case_budg_total_housing_grant_issuance = trim(mfip_case_budg_total_housing_grant_issuance)

			' Call write_value_and_transmit("X", 15, 3)			'State food benefit pop-up - I think this is duplicate
			Call write_value_and_transmit("X", 9, 44)			'Supplement pop-up
			EMReadScreen mfip_case_budg_food_supplement, 		10, 11, 32
			EMReadScreen mfip_case_budg_state_food_supplement, 	10, 16, 32
			EMReadScreen mfip_case_budg_cash_supplement, 		10, 11, 68
			EMReadScreen mfip_case_budg_housing_grant_supplement, 10, 16, 68

			mfip_case_budg_food_supplement = trim(mfip_case_budg_food_supplement)
			mfip_case_budg_state_food_supplement = trim(mfip_case_budg_state_food_supplement)
			mfip_case_budg_cash_supplement = trim(mfip_case_budg_cash_supplement)
			mfip_case_budg_housing_grant_supplement = trim(mfip_case_budg_housing_grant_supplement)
			transmit

			' Call write_value_and_transmit("X", 10, 44)			'Overpayment pop-up - MAYBE WE DON"T NEED THIS?
			Call write_value_and_transmit("X", 12, 44)			'Recoupment pop-up
			EMReadScreen mfip_case_budg_cash_recoupment, 10, 7, 51
			EMReadScreen mfip_case_budg_state_food_recoupment, 10, 8, 51
			EMReadScreen mfip_case_budg_food_recoupment, 10, 9, 51

			mfip_case_budg_cash_recoupment = trim(mfip_case_budg_cash_recoupment)
			mfip_case_budg_state_food_recoupment = trim(mfip_case_budg_state_food_recoupment)
			mfip_case_budg_food_recoupment = trim(mfip_case_budg_food_recoupment)
			transmit

			Call write_value_and_transmit("X", 14, 44)			'Total Food issuance pop-up
			EMReadScreen mfip_case_budg_fed_food_memb_count, 1, 7, 17
			EMReadScreen mfip_case_budg_fed_food_benefit_amt, 10, 7, 45
			EMReadScreen mfip_case_budg_state_food_memb_count, 1, 9, 17
			EMReadScreen mfip_case_budg_state_food_benefit_amt, 10, 9, 45

			mfip_case_budg_fed_food_memb_count = trim(mfip_case_budg_fed_food_memb_count)
			mfip_case_budg_fed_food_benefit_amt = trim(mfip_case_budg_fed_food_benefit_amt)
			mfip_case_budg_state_food_memb_count = trim(mfip_case_budg_state_food_memb_count)
			mfip_case_budg_state_food_benefit_amt = trim(mfip_case_budg_state_food_benefit_amt)
			transmit

			Call write_value_and_transmit("X", 15, 44)			'Total Cash Issuance pop-up
			EMReadScreen mfip_case_budg_tanf_cash_memb_count, 1, 8, 17
			EMReadScreen mfip_case_budg_tanf_cash_benefit_amt, 10, 8, 45
			EMReadScreen mfip_case_budg_state_cash_memb_count, 1, 10, 17
			EMReadScreen mfip_case_budg_state_cash_benefit_amt, 10, 10, 45

			mfip_case_budg_tanf_cash_memb_count = trim(mfip_case_budg_tanf_cash_memb_count)
			mfip_case_budg_tanf_cash_benefit_amt = trim(mfip_case_budg_tanf_cash_benefit_amt)
			mfip_case_budg_state_cash_memb_count = trim(mfip_case_budg_state_cash_memb_count)
			mfip_case_budg_state_cash_benefit_amt = trim(mfip_case_budg_state_cash_benefit_amt)
			transmit
			' Call write_value_and_transmit("X", 16, 44)			'MFIP Housing Grant Issuance pop-up - there is not federal housing grant
			transmit			'MFSM
			' MsgBox "In MFSM"
			EMReadScreen mfip_approved_date, 8, 3, 14
			EMReadScreen mfip_process_date, 8, 2, 73
			EMReadScreen mfip_prev_approval, 4, 3, 73

			EMReadScreen mfip_case_last_approval_date, 8, 5, 31
			EMReadScreen mfip_case_current_prog_status, 12, 6, 31
			EMReadScreen mfip_case_eligibility_result, 12,  7, 31
			EMReadScreen mfip_case_hrf_reporting, 12, 8, 31
			EMReadScreen mfip_case_source_of_info, 4, 9, 31
			EMReadScreen mfip_case_benefit_impact, 12, 10, 31
			EMReadScreen mfip_case_review_date, 8, 11, 31
			EMReadScreen mfip_case_budget_cycle, 12, 12, 31
			EMReadScreen mfip_case_vendor_reason_code, 2, 13, 31

			EMReadScreen mfip_case_responsible_county, 2, 5, 73
			EMReadScreen mfip_case_service_county, 2, 6, 73
			EMReadScreen mfip_case_asst_unit_caregivers, 1, 7, 73
			EMReadScreen mfip_case_asst_unit_children, 2, 8, 73
			EMReadScreen mfip_case_total_assets, 10, 9, 71
			EMReadScreen mfip_case_maximum_assets, 10, 10, 71
			EMReadScreen mfip_case_summary_grant_amount, 10, 11, 71
			EMReadScreen mfip_case_summary_net_grant_amount, 10, 13, 71
			EMReadScreen mfip_case_summary_cash_portion, 10, 14, 71
			EMReadScreen mfip_case_summary_food_portion, 10, 15, 71
			EMReadScreen mfip_case_summary_housing_grant, 10, 16, 71

			If mfip_case_vendor_reason_code = "01" Then mfip_case_vendor_reason_info = "Client Request"
			If mfip_case_vendor_reason_code = "05" Then mfip_case_vendor_reason_info = "Money Mismanagement"
			If mfip_case_vendor_reason_code = "06" Then mfip_case_vendor_reason_info = "Social Service Non-Coop"
			If mfip_case_vendor_reason_code = "07" Then mfip_case_vendor_reason_info = "Residing in a Facility"
			If mfip_case_vendor_reason_code = "21" Then mfip_case_vendor_reason_info = "MFIP Sanction Related Vendor"
			If mfip_case_vendor_reason_code = "22" Then mfip_case_vendor_reason_info = "Convicted Drug Felon in Household"

			mfip_prev_approval = trim(mfip_prev_approval)
			mfip_case_last_approval_date = trim(mfip_case_last_approval_date)

			mfip_case_current_prog_status = trim(mfip_case_current_prog_status)
			mfip_case_eligibility_result = trim(mfip_case_eligibility_result)
			mfip_case_hrf_reporting = trim(mfip_case_hrf_reporting)
			mfip_case_source_of_info = trim(mfip_case_source_of_info)
			mfip_case_benefit_impact = trim(mfip_case_benefit_impact)

			mfip_case_budget_cycle = trim(mfip_case_budget_cycle)
			mfip_case_vendor_reason_code = trim(mfip_case_vendor_reason_code)

			mfip_case_asst_unit_caregivers = trim(mfip_case_asst_unit_caregivers)
			mfip_case_asst_unit_children = trim(mfip_case_asst_unit_children)
			mfip_case_total_assets = trim(mfip_case_total_assets)
			mfip_case_maximum_assets = trim(mfip_case_maximum_assets)
			mfip_case_summary_grant_amount = trim(mfip_case_summary_grant_amount)
			mfip_case_summary_net_grant_amount = trim(mfip_case_summary_net_grant_amount)
			mfip_case_summary_cash_portion = trim(mfip_case_summary_cash_portion)
			mfip_case_summary_food_portion = trim(mfip_case_summary_food_portion)
			mfip_case_summary_housing_grant = trim(mfip_case_summary_housing_grant)
			' Msgbox mfip_case_summary_net_grant_amount

			If mfip_case_asst_unit_caregivers = "0" Then mfip_child_only = True
		End If

		Call Back_to_SELF
	end sub

end class

class msa_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found
	public er_month
	public er_status
	public er_caf_date
	public er_interview_date

	public msa_elig_ref_numbs()
	public msa_elig_membs_full_name()
	public msa_elig_membs_request_yn()
	public msa_elig_membs_member_code()
	public msa_elig_membs_member_info()
	public msa_elig_membs_elig_status()
	public msa_elig_membs_elig_basis_code()
	public msa_elig_membs_elig_basis_info()
	public msa_elig_membs_begin_date()
	public msa_elig_membs_budget_cycle()
	public msa_elig_membs_test_absence()
	public msa_elig_membs_test_age()
	public msa_elig_membs_test_basis_of_eligibility()
	public msa_elig_membs_test_citizenship()
	public msa_elig_membs_test_dupl_assistance()
	public msa_elig_membs_test_fail_coop()
	public msa_elig_membs_test_fraud()
	public msa_elig_membs_test_ive_eligible()
	public msa_elig_membs_test_living_arrangement()
	public msa_elig_membs_test_ssi_basis()
	public msa_elig_membs_test_ssn_coop()
	public msa_elig_membs_test_unit_member()
	public msa_elig_membs_test_verif()
	public msa_elig_membs_test_absence_absent()
	public msa_elig_membs_test_absence_death()
	public msa_elig_membs_test_fail_coop_sign_iaas()
	public msa_elig_membs_test_fail_coop_applied_other_benefits()
	public msa_elig_membs_test_unit_member_faci()
	public msa_elig_membs_test_unit_member_relationship()
	public msa_elig_membs_test_verif_date_of_birth()
	public msa_elig_membs_test_verif_disability()
	public msa_elig_membs_test_verif_ssi()

	public msa_elig_budg_memb_gross_earned_income()
	public msa_elig_budg_memb_blind_disa_student()
	public msa_elig_budg_memb_standard_disregard()
	public msa_elig_budg_memb_earned_income()
	public msa_elig_budg_memb_standard_EI_disregard()
	public msa_elig_budg_memb_work_expense_disa()
	public msa_elig_budg_memb_earned_inc_subtotal()
	public msa_elig_budg_memb_earned_inc_disregard()
	public msa_elig_budg_memb_work_expense_blind()
	public msa_elig_budg_memb_net_earned_income()
	public msa_elig_budg_memb_special_needs_total()

	public msa_elig_case_test_applicant_eligible
	public msa_elig_case_test_application_withdrawn
	public msa_elig_case_test_eligible_member
	public msa_elig_case_test_fail_file
	public msa_elig_case_test_prosp_gross_income
	public msa_elig_case_test_prosp_net_income
	public msa_elig_case_test_residence
	public msa_elig_case_test_assets
	public msa_elig_case_test_retro_net_income
	public msa_elig_case_test_verif
	public msa_elig_case_shared_hh_yn

	public msa_elig_case_test_fail_file_revw
	public msa_elig_case_test_fail_file_hrf
	public msa_elig_case_test_prosp_gross_earned_income
	public msa_elig_case_test_prosp_gross_unearned_income
	public msa_elig_case_test_prosp_gross_deemed_income
	public msa_elig_case_test_prosp_total_gross_income
	public msa_elig_case_test_prosp_gross_ssi_need_standard
	public msa_elig_case_test_prosp_gross_ssi_standard_multiplier
	public msa_elig_case_test_prosp_gross_income_limit
	public msa_elig_case_test_total_countable_assets
	public msa_elig_case_test_maximum_assets
	public msa_elig_case_test_verif_acct
	public msa_elig_case_test_verif_addr
	public msa_elig_case_test_verif_busi
	public msa_elig_case_test_verif_cars
	public msa_elig_case_test_verif_jobs
	public msa_elig_case_test_verif_lump
	public msa_elig_case_test_verif_pact
	public msa_elig_case_test_verif_rbic
	public msa_elig_case_test_verif_secu
	public msa_elig_case_test_verif_spon
	public msa_elig_case_test_verif_stin
	public msa_elig_case_test_verif_unea

	public msa_elig_case_budg_type

	public msa_elig_budg_ssi_standard_fbr
	public msa_elig_budg_standard_disregard
	public msa_elig_budg_unearned_income
	public msa_elig_budg_deemed_income
	public msa_elig_budg_net_unearned_income
	public msa_elig_budg_net_earned_income

	public msa_elig_budg_spec_standard_ref_numb()
	public msa_elig_budg_spec_standard_type_code()
	public msa_elig_budg_spec_standard_type_info()
	public msa_elig_budg_spec_standard_amount()

	public msa_elig_budg_need_standard
	public msa_elig_budg_net_income
	public msa_elig_budg_msa_grant
	public msa_elig_budg_amount_already_issued
	public msa_elig_budg_supplement_due
	public msa_elig_budg_overpayment
	public msa_elig_budg_adjusted_grant_amount
	public msa_elig_budg_recoupment
	public msa_elig_budg_current_payment

	public msa_elig_budg_basic_needs_assistance_standard
	public msa_elig_budg_special_needs
	public msa_elig_budg_household_total_needs

	public msa_elig_summ_approved_date
	public msa_elig_summ_process_date
	public msa_elig_summ_date_last_approval
	public msa_elig_summ_curr_prog_status
	public msa_elig_summ_eligibility_result
	public msa_elig_summ_reporting_status
	public msa_elig_summ_source_of_info
	public msa_elig_summ_benefit
	public msa_elig_summ_recertification_date
	public msa_elig_summ_budget_cycle
	public msa_elig_summ_eligible_houshold_members
	public msa_elig_summ_shared_houshold
	public msa_elig_summ_vendor_reason_code
	public msa_elig_summ_vendor_reason_info
	public msa_elig_summ_responsible_county
	public msa_elig_summ_servicing_county
	public msa_elig_summ_total_assets
	public msa_elig_summ_maximum_assets
	public msa_elig_summ_grant
	public msa_elig_summ_current_payment
	public msa_elig_summ_worker_message

	public sub read_elig()
		approved_today = False
		approved_version_found = False

		call navigate_to_MAXIS_screen("ELIG", "MSA ")
		EMWriteScreen elig_footer_month, 20, 56
		EMWriteScreen elig_footer_year, 20, 59
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

			ReDim msa_elig_ref_numbs(0)
			ReDim msa_elig_membs_full_name(0)
			ReDim msa_elig_membs_request_yn(0)
			ReDim msa_elig_membs_member_code(0)
			ReDim msa_elig_membs_member_info(0)
			ReDim msa_elig_membs_elig_status(0)
			ReDim msa_elig_membs_elig_basis_code(0)
			ReDim msa_elig_membs_elig_basis_info(0)
			ReDim msa_elig_membs_begin_date(0)
			ReDim msa_elig_membs_budget_cycle(0)
			ReDim msa_elig_membs_test_absence(0)
			ReDim msa_elig_membs_test_age(0)
			ReDim msa_elig_membs_test_basis_of_eligibility(0)
			ReDim msa_elig_membs_test_citizenship(0)
			ReDim msa_elig_membs_test_dupl_assistance(0)
			ReDim msa_elig_membs_test_fail_coop(0)
			ReDim msa_elig_membs_test_fraud(0)
			ReDim msa_elig_membs_test_ive_eligible(0)
			ReDim msa_elig_membs_test_living_arrangement(0)
			ReDim msa_elig_membs_test_ssi_basis(0)
			ReDim msa_elig_membs_test_ssn_coop(0)
			ReDim msa_elig_membs_test_unit_member(0)
			ReDim msa_elig_membs_test_verif(0)
			ReDim msa_elig_membs_test_absence_absent(0)
			ReDim msa_elig_membs_test_absence_death(0)
			ReDim msa_elig_membs_test_fail_coop_sign_iaas(0)
			ReDim msa_elig_membs_test_fail_coop_applied_other_benefits(0)
			ReDim msa_elig_membs_test_unit_member_faci(0)
			ReDim msa_elig_membs_test_unit_member_relationship(0)
			ReDim msa_elig_membs_test_verif_date_of_birth(0)
			ReDim msa_elig_membs_test_verif_disability(0)
			ReDim msa_elig_membs_test_verif_ssi(0)
			ReDim msa_elig_budg_memb_gross_earned_income(0)
			ReDim msa_elig_budg_memb_blind_disa_student(0)
			ReDim msa_elig_budg_memb_standard_disregard(0)
			ReDim msa_elig_budg_memb_earned_income(0)
			ReDim msa_elig_budg_memb_standard_EI_disregard(0)
			ReDim msa_elig_budg_memb_work_expense_disa(0)
			ReDim msa_elig_budg_memb_earned_inc_subtotal(0)
			ReDim msa_elig_budg_memb_earned_inc_disregard(0)
			ReDim msa_elig_budg_memb_work_expense_blind(0)
			ReDim msa_elig_budg_memb_net_earned_income(0)
			ReDim msa_elig_budg_memb_special_needs_total(0)


			ReDim msa_elig_budg_spec_standard_ref_numb(0)
			ReDim msa_elig_budg_spec_standard_type_code(0)
			ReDim msa_elig_budg_spec_standard_type_info(0)
			ReDim msa_elig_budg_spec_standard_amount(0)

			elig_memb_count = 0
			msa_row = 7
			Do
				EMReadScreen ref_numb, 2, msa_row, 5

				ReDim preserve msa_elig_ref_numbs(elig_memb_count)
				ReDim preserve msa_elig_membs_full_name(elig_memb_count)
				ReDim preserve msa_elig_membs_request_yn(elig_memb_count)
				ReDim preserve msa_elig_membs_member_code(elig_memb_count)
				ReDim preserve msa_elig_membs_member_info(elig_memb_count)
				ReDim preserve msa_elig_membs_elig_status(elig_memb_count)
				ReDim preserve msa_elig_membs_elig_basis_code(elig_memb_count)
				ReDim preserve msa_elig_membs_elig_basis_info(elig_memb_count)
				ReDim preserve msa_elig_membs_begin_date(elig_memb_count)
				ReDim preserve msa_elig_membs_budget_cycle(elig_memb_count)
				ReDim preserve msa_elig_membs_test_absence(elig_memb_count)
				ReDim preserve msa_elig_membs_test_age(elig_memb_count)
				ReDim preserve msa_elig_membs_test_basis_of_eligibility(elig_memb_count)
				ReDim preserve msa_elig_membs_test_citizenship(elig_memb_count)
				ReDim preserve msa_elig_membs_test_dupl_assistance(elig_memb_count)
				ReDim preserve msa_elig_membs_test_fail_coop(elig_memb_count)
				ReDim preserve msa_elig_membs_test_fraud(elig_memb_count)
				ReDim preserve msa_elig_membs_test_ive_eligible(elig_memb_count)
				ReDim preserve msa_elig_membs_test_living_arrangement(elig_memb_count)
				ReDim preserve msa_elig_membs_test_ssi_basis(elig_memb_count)
				ReDim preserve msa_elig_membs_test_ssn_coop(elig_memb_count)
				ReDim preserve msa_elig_membs_test_unit_member(elig_memb_count)
				ReDim preserve msa_elig_membs_test_verif(elig_memb_count)
				ReDim preserve msa_elig_membs_test_absence_absent(elig_memb_count)
				ReDim preserve msa_elig_membs_test_absence_death(elig_memb_count)
				ReDim preserve msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count)
				ReDim preserve msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count)
				ReDim preserve msa_elig_membs_test_unit_member_faci(elig_memb_count)
				ReDim preserve msa_elig_membs_test_unit_member_relationship(elig_memb_count)
				ReDim preserve msa_elig_membs_test_verif_date_of_birth(elig_memb_count)
				ReDim preserve msa_elig_membs_test_verif_disability(elig_memb_count)
				ReDim preserve msa_elig_membs_test_verif_ssi(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_gross_earned_income(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_blind_disa_student(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_standard_disregard(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_earned_income(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_standard_EI_disregard(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_work_expense_disa(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_earned_inc_subtotal(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_earned_inc_disregard(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_work_expense_blind(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_net_earned_income(elig_memb_count)
				ReDim preserve msa_elig_budg_memb_special_needs_total(elig_memb_count)

				msa_elig_ref_numbs(elig_memb_count) = ref_numb

				EMReadScreen msa_elig_membs_request_yn(elig_memb_count), 1, msa_row, 25

				EMReadScreen msa_elig_membs_member_code(elig_memb_count), 1, msa_row, 29
				If msa_elig_membs_member_code(elig_memb_count) = "A" Then msa_elig_membs_member_info(elig_memb_count) = "Eligible"
				If msa_elig_membs_member_code(elig_memb_count) = "1" Then msa_elig_membs_member_info(elig_memb_count) = "Non-MSA Spouse"
				If msa_elig_membs_member_code(elig_memb_count) = "2" Then msa_elig_membs_member_info(elig_memb_count) = "Non-MSA Parent - Deem Income/Resources"
				If msa_elig_membs_member_code(elig_memb_count) = "4" Then msa_elig_membs_member_info(elig_memb_count) = "Step Parent - Deem Resources"
				If msa_elig_membs_member_code(elig_memb_count) = "N" Then msa_elig_membs_member_info(elig_memb_count) = "Not Counted"
				If msa_elig_membs_member_code(elig_memb_count) = "I" Then msa_elig_membs_member_info(elig_memb_count) = "Ineligible"

				EMReadScreen msa_elig_membs_elig_status(elig_memb_count), 10, msa_row, 46
				msa_elig_membs_elig_status(elig_memb_count) = trim(msa_elig_membs_elig_status(elig_memb_count))

				EMReadScreen msa_elig_membs_elig_basis_code(elig_memb_count), 1, msa_row, 59
				If msa_elig_membs_elig_basis_code(elig_memb_count) = "A" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Aged"
				If msa_elig_membs_elig_basis_code(elig_memb_count) = "B" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Blind"
				If msa_elig_membs_elig_basis_code(elig_memb_count) = "D" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Disabled"
				If msa_elig_membs_elig_basis_code(elig_memb_count) = "S" Then msa_elig_membs_elig_basis_info(elig_memb_count) = "SSI"
				If msa_elig_membs_elig_basis_code(elig_memb_count) = " " Then msa_elig_membs_elig_basis_info(elig_memb_count) = "Blank"

				EMReadScreen msa_elig_membs_begin_date(elig_memb_count), 8, msa_row, 63
				msa_elig_membs_begin_date(elig_memb_count) = trim(msa_elig_membs_begin_date(elig_memb_count))
				If msa_elig_membs_begin_date(elig_memb_count) <> "" then msa_elig_membs_begin_date(elig_memb_count) = replace(msa_elig_membs_begin_date(elig_memb_count), " ", "/")

				EMReadScreen msa_elig_membs_budget_cycle(elig_memb_count), 1, msa_row, 76
				If msa_elig_membs_budget_cycle(elig_memb_count) = "P" Then msa_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
				If msa_elig_membs_budget_cycle(elig_memb_count) = "R" Then msa_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

				Call write_value_and_transmit("X", msa_row, 3)

				EMReadScreen full_name_information, 20, 7, 10
				full_name_information = trim(full_name_information)
				name_array = split(full_name_information, " ")
				For each name_parts in name_array
					If len(name_parts) <> 1 Then msa_elig_membs_full_name(elig_memb_count) = msa_elig_membs_full_name(elig_memb_count) & " " & name_parts
				Next
				msa_elig_membs_full_name(elig_memb_count) = trim((msa_elig_membs_full_name(elig_memb_count)))

				EMReadScreen msa_elig_membs_test_absence(elig_memb_count), 				6, 10, 8
				EMReadScreen msa_elig_membs_test_age(elig_memb_count), 					6, 11, 8
				EMReadScreen msa_elig_membs_test_basis_of_eligibility(elig_memb_count), 6, 12, 8
				EMReadScreen msa_elig_membs_test_citizenship(elig_memb_count), 			6, 13, 8
				EMReadScreen msa_elig_membs_test_dupl_assistance(elig_memb_count), 		6, 14, 8
				EMReadScreen msa_elig_membs_test_fail_coop(elig_memb_count), 			6, 15, 8
				EMReadScreen msa_elig_membs_test_fraud(elig_memb_count), 				6, 16, 8

				EMReadScreen msa_elig_membs_test_ive_eligible(elig_memb_count), 		6, 10, 47
				EMReadScreen msa_elig_membs_test_living_arrangement(elig_memb_count), 	6, 11, 47
				EMReadScreen msa_elig_membs_test_ssi_basis(elig_memb_count), 			6, 12, 47
				EMReadScreen msa_elig_membs_test_ssn_coop(elig_memb_count), 			6, 13, 47
				EMReadScreen msa_elig_membs_test_unit_member(elig_memb_count), 			6, 14, 47
				EMReadScreen msa_elig_membs_test_verif(elig_memb_count), 				6, 15, 47

				msa_elig_membs_test_absence(elig_memb_count) = trim(msa_elig_membs_test_absence(elig_memb_count))
				msa_elig_membs_test_age(elig_memb_count) = trim(msa_elig_membs_test_age(elig_memb_count))
				msa_elig_membs_test_basis_of_eligibility(elig_memb_count) = trim(msa_elig_membs_test_basis_of_eligibility(elig_memb_count))
				msa_elig_membs_test_citizenship(elig_memb_count) = trim(msa_elig_membs_test_citizenship(elig_memb_count))
				msa_elig_membs_test_dupl_assistance(elig_memb_count) = trim(msa_elig_membs_test_dupl_assistance(elig_memb_count))
				msa_elig_membs_test_fail_coop(elig_memb_count) = trim(msa_elig_membs_test_fail_coop(elig_memb_count))
				msa_elig_membs_test_fraud(elig_memb_count) = trim(msa_elig_membs_test_fraud(elig_memb_count))

				msa_elig_membs_test_ive_eligible(elig_memb_count) = trim(msa_elig_membs_test_ive_eligible(elig_memb_count))
				msa_elig_membs_test_living_arrangement(elig_memb_count) = trim(msa_elig_membs_test_living_arrangement(elig_memb_count))
				msa_elig_membs_test_ssi_basis(elig_memb_count) = trim(msa_elig_membs_test_ssi_basis(elig_memb_count))
				msa_elig_membs_test_ssn_coop(elig_memb_count) = trim(msa_elig_membs_test_ssn_coop(elig_memb_count))
				msa_elig_membs_test_unit_member(elig_memb_count) = trim(msa_elig_membs_test_unit_member(elig_memb_count))
				msa_elig_membs_test_verif(elig_memb_count) = trim(msa_elig_membs_test_verif(elig_memb_count))

				Call write_value_and_transmit("X", 10, 6)
				EMReadScreen msa_elig_membs_test_absence_absent(elig_memb_count), 	6, 12, 40
				EMReadScreen msa_elig_membs_test_absence_death(elig_memb_count), 	6, 13, 40

				msa_elig_membs_test_absence_absent(elig_memb_count) = trim(msa_elig_membs_test_absence_absent(elig_memb_count))
				msa_elig_membs_test_absence_death(elig_memb_count) = trim(msa_elig_membs_test_absence_death(elig_memb_count))
				transmit

				Call write_value_and_transmit("X", 15, 6)
				EMReadScreen msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count), 				6, 12, 24
				EMReadScreen msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count), 6, 13, 24

				msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count) = trim(msa_elig_membs_test_fail_coop_sign_iaas(elig_memb_count))
				msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count) = trim(msa_elig_membs_test_fail_coop_applied_other_benefits(elig_memb_count))
				transmit

				Call write_value_and_transmit("X", 14, 45)
				EMReadScreen msa_elig_membs_test_unit_member_faci(elig_memb_count), 		6, 12, 24
				EMReadScreen msa_elig_membs_test_unit_member_relationship(elig_memb_count), 6, 13, 24

				msa_elig_membs_test_unit_member_faci(elig_memb_count) = trim(msa_elig_membs_test_unit_member_faci(elig_memb_count))
				msa_elig_membs_test_unit_member_relationship(elig_memb_count) = trim(msa_elig_membs_test_unit_member_relationship(elig_memb_count))
				transmit

				Call write_value_and_transmit("X", 15, 45)
				EMReadScreen msa_elig_membs_test_verif_date_of_birth(elig_memb_count), 	6, 12, 30
				EMReadScreen msa_elig_membs_test_verif_disability(elig_memb_count), 	6, 13, 30
				EMReadScreen msa_elig_membs_test_verif_ssi(elig_memb_count), 			6, 14, 30

				msa_elig_membs_test_verif_date_of_birth(elig_memb_count) = trim(msa_elig_membs_test_verif_date_of_birth(elig_memb_count))
				msa_elig_membs_test_verif_disability(elig_memb_count) = trim(msa_elig_membs_test_verif_disability(elig_memb_count))
				msa_elig_membs_test_verif_ssi(elig_memb_count) = trim(msa_elig_membs_test_verif_ssi(elig_memb_count))
				transmit

				transmit

				msa_row = msa_row + 1
				elig_memb_count = elig_memb_count + 1
				EMReadScreen next_ref_numb, 2, msa_row, 5
			Loop until next_ref_numb = "  "
			transmit 		'going to the next panel - MSCR

			EMReadScreen msa_elig_case_test_applicant_eligible, 	6, 6, 7
			EMReadScreen msa_elig_case_test_application_withdrawn, 	6, 7, 7
			EMReadScreen msa_elig_case_test_eligible_member, 		6, 8, 7
			EMReadScreen msa_elig_case_test_fail_file, 				6, 9, 7
			EMReadScreen msa_elig_case_test_prosp_gross_income, 	6, 10, 7

			EMReadScreen msa_elig_case_test_prosp_net_income, 	6, 6, 45
			EMReadScreen msa_elig_case_test_residence, 			6, 7, 45
			EMReadScreen msa_elig_case_test_assets, 			6, 8, 45
			EMReadScreen msa_elig_case_test_retro_net_income, 	6, 9, 45
			EMReadScreen msa_elig_case_test_verif, 				6, 10, 45

			EMReadScreen msa_elig_case_shared_hh_yn, 1, 13, 61

			msa_elig_case_test_applicant_eligible = trim(msa_elig_case_test_applicant_eligible)
			msa_elig_case_test_application_withdrawn = trim(msa_elig_case_test_application_withdrawn)
			msa_elig_case_test_eligible_member = trim(msa_elig_case_test_eligible_member)
			msa_elig_case_test_fail_file = trim(msa_elig_case_test_fail_file)
			msa_elig_case_test_prosp_gross_income = trim(msa_elig_case_test_prosp_gross_income)

			msa_elig_case_test_prosp_net_income = trim(msa_elig_case_test_prosp_net_income)
			msa_elig_case_test_residence = trim(msa_elig_case_test_residence)
			msa_elig_case_test_assets = trim(msa_elig_case_test_assets)
			msa_elig_case_test_retro_net_income = trim(msa_elig_case_test_retro_net_income)
			msa_elig_case_test_verif = trim(msa_elig_case_test_verif)

			If msa_elig_case_test_fail_file <> "NA" Then
				Call write_value_and_transmit("X", 9, 5)

				EMReadScreen msa_elig_case_test_fail_file_revw, 6, 8, 28
				EMReadScreen msa_elig_case_test_fail_file_hrf, 6, 9, 28

				msa_elig_case_test_fail_file_revw = trim(msa_elig_case_test_fail_file_revw)
				msa_elig_case_test_fail_file_hrf = trim(msa_elig_case_test_fail_file_hrf)
				transmit
			End If

			If msa_elig_case_test_prosp_gross_income <> "NA" Then
				Call write_value_and_transmit("X", 10, 5)

				EMReadScreen msa_elig_case_test_prosp_gross_earned_income, 		9, 9, 55
				EMReadScreen msa_elig_case_test_prosp_gross_unearned_income, 	9, 10, 55
				EMReadScreen msa_elig_case_test_prosp_gross_deemed_income, 		9, 11, 55

				EMReadScreen msa_elig_case_test_prosp_total_gross_income, 			9, 13, 55
				EMReadScreen msa_elig_case_test_prosp_gross_ssi_need_standard, 		9, 14, 55
				EMReadScreen msa_elig_case_test_prosp_gross_ssi_standard_multiplier, 1, 15, 63
				EMReadScreen msa_elig_case_test_prosp_gross_income_limit, 			9, 16, 55


				msa_elig_case_test_prosp_gross_earned_income = trim(msa_elig_case_test_prosp_gross_earned_income)
				msa_elig_case_test_prosp_gross_unearned_income = trim(msa_elig_case_test_prosp_gross_unearned_income)
				msa_elig_case_test_prosp_gross_deemed_income = trim(msa_elig_case_test_prosp_gross_deemed_income)

				msa_elig_case_test_prosp_total_gross_income = trim(msa_elig_case_test_prosp_total_gross_income)
				msa_elig_case_test_prosp_gross_ssi_need_standard = trim(msa_elig_case_test_prosp_gross_ssi_need_standard)
				msa_elig_case_test_prosp_gross_income_limit = trim(msa_elig_case_test_prosp_gross_income_limit)
				transmit
			End If

			If msa_elig_case_test_assets <> "NA" Then
				Call write_value_and_transmit("X", 8, 43)

				EMReadScreen msa_elig_case_test_total_countable_assets, 10, 8, 48
				EMReadScreen msa_elig_case_test_maximum_assets, 		10, 9, 48

				msa_elig_case_test_total_countable_assets = replace(msa_elig_case_test_total_countable_assets, "_", "")
				msa_elig_case_test_maximum_assets = replace(msa_elig_case_test_maximum_assets, "_", "")
				transmit
			End If

			If msa_elig_case_test_verif <> "NA" Then
				Call write_value_and_transmit("X", 10, 43)

				EMReadScreen msa_elig_case_test_verif_acct, 6, 6, 32
				EMReadScreen msa_elig_case_test_verif_addr, 6, 7, 32
				EMReadScreen msa_elig_case_test_verif_busi, 6, 8, 32
				EMReadScreen msa_elig_case_test_verif_cars, 6, 9, 32
				EMReadScreen msa_elig_case_test_verif_jobs, 6, 10, 32
				EMReadScreen msa_elig_case_test_verif_lump, 6, 11, 32
				EMReadScreen msa_elig_case_test_verif_pact, 6, 12, 32
				EMReadScreen msa_elig_case_test_verif_rbic, 6, 13, 32
				EMReadScreen msa_elig_case_test_verif_secu, 6, 14, 32
				EMReadScreen msa_elig_case_test_verif_spon, 6, 15, 32
				EMReadScreen msa_elig_case_test_verif_stin, 6, 16, 32
				EMReadScreen msa_elig_case_test_verif_unea, 6, 17, 32

				msa_elig_case_test_verif_acct = trim(msa_elig_case_test_verif_acct)
				msa_elig_case_test_verif_addr = trim(msa_elig_case_test_verif_addr)
				msa_elig_case_test_verif_busi = trim(msa_elig_case_test_verif_busi)
				msa_elig_case_test_verif_cars = trim(msa_elig_case_test_verif_cars)
				msa_elig_case_test_verif_jobs = trim(msa_elig_case_test_verif_jobs)
				msa_elig_case_test_verif_lump = trim(msa_elig_case_test_verif_lump)
				msa_elig_case_test_verif_pact = trim(msa_elig_case_test_verif_pact)
				msa_elig_case_test_verif_rbic = trim(msa_elig_case_test_verif_rbic)
				msa_elig_case_test_verif_secu = trim(msa_elig_case_test_verif_secu)
				msa_elig_case_test_verif_spon = trim(msa_elig_case_test_verif_spon)
				msa_elig_case_test_verif_stin = trim(msa_elig_case_test_verif_stin)
				msa_elig_case_test_verif_unea = trim(msa_elig_case_test_verif_unea)
				transmit
			End If

			transmit 		'going to the next panel - MSCB

			EmReadScreen msa_elig_case_budg_type, 12, 3, 25
			msa_elig_case_budg_type = trim(msa_elig_case_budg_type)

			If msa_elig_case_budg_type = "SSI TYPE" Then
				EMReadScreen msa_elig_budg_ssi_standard_fbr, 	9, 6, 32
				EMReadScreen msa_elig_budg_standard_disregard, 	9, 7, 32

				msa_elig_budg_ssi_standard_fbr = trim(msa_elig_budg_ssi_standard_fbr)
				msa_elig_budg_standard_disregard = trim(msa_elig_budg_standard_disregard)
			End If

			If msa_elig_case_budg_type = "Non-SSI TYPE" Then
				EMReadScreen msa_elig_budg_unearned_income, 	9, 6, 32
				EMReadScreen msa_elig_budg_deemed_income, 		9, 7, 32
				EMReadScreen msa_elig_budg_standard_disregard, 	9, 8, 32
				EMReadScreen msa_elig_budg_net_unearned_income, 9, 9, 32
				EMReadScreen msa_elig_budg_net_earned_income, 	9, 10, 32

				msa_elig_budg_unearned_income = trim(msa_elig_budg_unearned_income)
				msa_elig_budg_deemed_income = trim(msa_elig_budg_deemed_income)
				msa_elig_budg_standard_disregard = trim(msa_elig_budg_standard_disregard)
				msa_elig_budg_net_unearned_income = trim(msa_elig_budg_net_unearned_income)
				msa_elig_budg_net_earned_income = trim(msa_elig_budg_net_earned_income)

				Call write_value_and_transmit("X", 10, 3)

				EMReadScreen msa_elig_budg_gross_earned_income, 	9, 9, 42
				EMReadScreen msa_elig_budg_blind_disa_student, 		9, 10, 42
				EMReadScreen msa_elig_budg_earned_standard_disregard, 9, 11, 42
				EMReadScreen msa_elig_budg_earned_income, 			9, 12, 42
				EMReadScreen msa_elig_budg_standard_EI_disregard, 	9, 13, 42
				EMReadScreen msa_elig_budg_work_expense_disa, 		9, 14, 42
				EMReadScreen msa_elig_budg_earned_inc_subtotal, 	9, 15, 42
				EMReadScreen msa_elig_budg_earned_inc_disregard, 	9, 16, 42
				EMReadScreen msa_elig_budg_work_expense_blind, 		9, 17, 42

				EMReadScreen ref_numb_one, 2, 7, 62
				If ref_numb_one <> "  " Then
					For memn_count = 0 to UBound(msa_elig_ref_numbs)
						If ref_numb_one = msa_elig_ref_numbs(memn_count) Then
							EMReadScreen msa_elig_budg_memb_gross_earned_income(memn_count), 	9, 9, 54
							EMReadScreen msa_elig_budg_memb_blind_disa_student(memn_count), 	9, 10, 54
							EMReadScreen msa_elig_budg_memb_standard_disregard(memn_count), 	9, 11, 54
							EMReadScreen msa_elig_budg_memb_earned_income(memn_count), 			9, 12, 54
							EMReadScreen msa_elig_budg_memb_standard_EI_disregard(memn_count), 	9, 13, 54
							EMReadScreen msa_elig_budg_memb_work_expense_disa(memn_count), 		9, 14, 54
							EMReadScreen msa_elig_budg_memb_earned_inc_subtotal(memn_count), 	9, 15, 54
							EMReadScreen msa_elig_budg_memb_earned_inc_disregard(memn_count), 	9, 16, 54
							EMReadScreen msa_elig_budg_memb_work_expense_blind(memn_count), 	9, 17, 54
							EMReadScreen msa_elig_budg_memb_net_earned_income(memn_count), 		9, 18, 54

							msa_elig_budg_memb_gross_earned_income(memn_count) = trim(msa_elig_budg_memb_gross_earned_income(memn_count))
							msa_elig_budg_memb_blind_disa_student(memn_count) = trim(msa_elig_budg_memb_blind_disa_student(memn_count))
							msa_elig_budg_memb_standard_disregard(memn_count) = trim(msa_elig_budg_memb_standard_disregard(memn_count))
							msa_elig_budg_memb_earned_income(memn_count) = trim(msa_elig_budg_memb_earned_income(memn_count))
							msa_elig_budg_memb_standard_EI_disregard(memn_count) = trim(msa_elig_budg_memb_standard_EI_disregard(memn_count))
							msa_elig_budg_memb_work_expense_disa(memn_count) = trim(msa_elig_budg_memb_work_expense_disa(memn_count))
							msa_elig_budg_memb_earned_inc_subtotal(memn_count) = trim(msa_elig_budg_memb_earned_inc_subtotal(memn_count))
							msa_elig_budg_memb_earned_inc_disregard(memn_count) = trim(msa_elig_budg_memb_earned_inc_disregard(memn_count))
							msa_elig_budg_memb_work_expense_blind(memn_count) = trim(msa_elig_budg_memb_work_expense_blind(memn_count))
							msa_elig_budg_memb_net_earned_income(memn_count) = trim(msa_elig_budg_memb_net_earned_income(memn_count))
						End If
					Next
				End if

				EMReadScreen ref_numb_two, 2, 7, 75
				If ref_numb_two <> "  " Then
					For memn_count = 0 to UBound(msa_elig_ref_numbs)
						If ref_numb_two = msa_elig_ref_numbs(memn_count) Then
							EMReadScreen msa_elig_budg_memb_gross_earned_income(memn_count), 	9, 9, 67
							EMReadScreen msa_elig_budg_memb_blind_disa_student(memn_count), 	9, 10, 67
							EMReadScreen msa_elig_budg_memb_standard_disregard(memn_count), 	9, 11, 67
							EMReadScreen msa_elig_budg_memb_earned_income(memn_count), 			9, 12, 67
							EMReadScreen msa_elig_budg_memb_standard_EI_disregard(memn_count), 	9, 13, 67
							EMReadScreen msa_elig_budg_memb_work_expense_disa(memn_count), 		9, 14, 67
							EMReadScreen msa_elig_budg_memb_earned_inc_subtotal(memn_count), 	9, 15, 67
							EMReadScreen msa_elig_budg_memb_earned_inc_disregard(memn_count), 	9, 16, 67
							EMReadScreen msa_elig_budg_memb_work_expense_blind(memn_count), 	9, 17, 67
							EMReadScreen msa_elig_budg_memb_net_earned_income(memn_count), 		9, 18, 67

							msa_elig_budg_memb_gross_earned_income(memn_count) = trim(msa_elig_budg_memb_gross_earned_income(memn_count))
							msa_elig_budg_memb_blind_disa_student(memn_count) = trim(msa_elig_budg_memb_blind_disa_student(memn_count))
							msa_elig_budg_memb_standard_disregard(memn_count) = trim(msa_elig_budg_memb_standard_disregard(memn_count))
							msa_elig_budg_memb_earned_income(memn_count) = trim(msa_elig_budg_memb_earned_income(memn_count))
							msa_elig_budg_memb_standard_EI_disregard(memn_count) = trim(msa_elig_budg_memb_standard_EI_disregard(memn_count))
							msa_elig_budg_memb_work_expense_disa(memn_count) = trim(msa_elig_budg_memb_work_expense_disa(memn_count))
							msa_elig_budg_memb_earned_inc_subtotal(memn_count) = trim(msa_elig_budg_memb_earned_inc_subtotal(memn_count))
							msa_elig_budg_memb_earned_inc_disregard(memn_count) = trim(msa_elig_budg_memb_earned_inc_disregard(memn_count))
							msa_elig_budg_memb_work_expense_blind(memn_count) = trim(msa_elig_budg_memb_work_expense_blind(memn_count))
							msa_elig_budg_memb_net_earned_income(memn_count) = trim(msa_elig_budg_memb_net_earned_income(memn_count))
						End If
					Next
				End if
				transmit
			End If

			EMReadScreen msa_elig_budg_need_standard, 			9, 6, 72
			EMReadScreen msa_elig_budg_net_income, 				9, 7, 72
			EMReadScreen msa_elig_budg_msa_grant, 				9, 8, 72

			EMReadScreen msa_elig_budg_amount_already_issued, 	9, 11, 72
			EMReadScreen msa_elig_budg_supplement_due, 			9, 12, 72
			EMReadScreen msa_elig_budg_overpayment, 			9, 13, 72

			EMReadScreen msa_elig_budg_adjusted_grant_amount, 	9, 15, 72
			EMReadScreen msa_elig_budg_recoupment, 				9, 16, 72
			EMReadScreen msa_elig_budg_current_payment, 		9, 17, 72

			msa_elig_budg_need_standard = trim(msa_elig_budg_need_standard)
			msa_elig_budg_net_income = trim(msa_elig_budg_net_income)
			msa_elig_budg_msa_grant = trim(msa_elig_budg_msa_grant)

			msa_elig_budg_amount_already_issued = trim(msa_elig_budg_amount_already_issued)
			msa_elig_budg_supplement_due = trim(msa_elig_budg_supplement_due)
			msa_elig_budg_overpayment = trim(msa_elig_budg_overpayment)

			msa_elig_budg_adjusted_grant_amount = trim(msa_elig_budg_adjusted_grant_amount)
			msa_elig_budg_recoupment = trim(msa_elig_budg_recoupment)
			msa_elig_budg_current_payment = trim(msa_elig_budg_current_payment)


			Call write_value_and_transmit("X", 6, 43)
			EMReadScreen msa_elig_budg_basic_needs_assistance_standard, 10, 16, 59
			EMReadScreen msa_elig_budg_special_needs, 					10, 17, 59
			EMReadScreen msa_elig_budg_household_total_needs, 			10, 18, 59

			msa_elig_budg_basic_needs_assistance_standard = trim(msa_elig_budg_basic_needs_assistance_standard)
			msa_elig_budg_special_needs = trim(msa_elig_budg_special_needs)
			msa_elig_budg_household_total_needs = trim(msa_elig_budg_household_total_needs)

			msa_col = 6
			spec_needs_count = 0
			For msa_col = 6 to 42 step 36
				EMReadScreen ref_numb, 2, 5, msa_col+9
				If ref_numb <> "  " Then
					For msa_membs = 0 to UBound(msa_elig_ref_numbs)
						If msa_elig_ref_numbs(msa_membs) = ref_numb Then
							EMReadScreen amount_total, 8, 15, msa_col+26
							msa_elig_budg_memb_special_needs_total(msa_membs) = amount_total
						End If
					Next

					EMReadScreen info_code, 2, 8, msa_col
					Do while info_code <> "__"
						ReDim preserve msa_elig_budg_spec_standard_ref_numb(spec_needs_count)
						ReDim preserve msa_elig_budg_spec_standard_type_code(spec_needs_count)
						ReDim preserve msa_elig_budg_spec_standard_type_info(spec_needs_count)
						ReDim preserve msa_elig_budg_spec_standard_amount(spec_needs_count)

						msa_elig_budg_spec_standard_ref_numb(spec_needs_count) = ref_numb
						msa_elig_budg_spec_standard_type_code(spec_needs_count) = info_code
						If info_code = "" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = ""
						If info_code = "01" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - High Protein > 79 Gr/Day"
						If info_code = "02" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Control Protein 40-60 GR/DAY"
						If info_code = "03" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Control Protein < 40 GR/DAY"
						If info_code = "04" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Low Cholesterol"
						If info_code = "05" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - High Residue"
						If info_code = "06" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Pregnancy and Lactation"
						If info_code = "07" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Gluten Free"
						If info_code = "08" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Lactose Free"
						If info_code = "09" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Anti Dumping"
						If info_code = "10" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Hypoglycemic"
						If info_code = "11" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "DIET - Ketogenic"
						If info_code = "RP" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Representative Payee"
						If info_code = "GF" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Guardianship Fee Max"
						If info_code = "SN" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Shelter Need"
						If info_code = "RM" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Restaurant Meals"
						If info_code = "EN" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Excess Need"
						If info_code = "OT" Then msa_elig_budg_spec_standard_type_info(spec_needs_count) = "Other Need"

						EMReadScreen msa_elig_budg_spec_standard_amount(spec_needs_count), 8, msa_row, msa_col+26
						msa_elig_budg_spec_standard_amount(spec_needs_count) = trim(msa_elig_budg_spec_standard_amount(spec_needs_count))

						msa_row = msa_row + 1
						If msa_row = 14 Then MsgBox "MORE THAN SIX?"
						spec_needs_count = spec_needs_count + 1
						EMReadScreen info_code, 2, msa_row, msa_col
					Loop
				End If
			Next
			transmit

			transmit 		'going to the next panel - MSSM

			EMReadScreen msa_elig_summ_approved_date, 8, 3, 14
			EMReadScreen msa_elig_summ_process_date, 8, 2, 72
			EMReadScreen msa_elig_summ_date_last_approval, 8, 5, 29
			EMReadScreen msa_elig_summ_curr_prog_status, 12, 6, 29
			EMReadScreen msa_elig_summ_eligibility_result, 12, 7, 29
			EMReadScreen msa_elig_summ_reporting_status, 12, 8, 29
			EMReadScreen msa_elig_summ_source_of_info, 4, 10, 29
			EMReadScreen msa_elig_summ_benefit, 12, 11, 29
			EMReadScreen msa_elig_summ_recertification_date, 8, 12, 29
			EMReadScreen msa_elig_summ_budget_cycle, 5, 13, 29
			EMReadScreen msa_elig_summ_eligible_houshold_members, 1, 14, 29
			EMReadScreen msa_elig_summ_shared_houshold, 3, 15, 29
			EMReadScreen msa_elig_summ_vendor_reason_code, 2, 16, 18

			EMReadScreen msa_elig_summ_responsible_county, 2, 5, 73
			EMReadScreen msa_elig_summ_servicing_county, 2, 6, 73
			EMReadScreen msa_elig_summ_total_assets, 9, 7, 72
			EMReadScreen msa_elig_summ_maximum_assets, 9, 8, 72
			EMReadScreen msa_elig_summ_grant, 9, 11, 72
			EMReadScreen msa_elig_summ_current_payment, 9, 17, 72

			EMReadScreen msa_elig_summ_worker_message, 80, 18, 1

			msa_elig_summ_curr_prog_status = trim(msa_elig_summ_curr_prog_status)
			msa_elig_summ_eligibility_result = trim(msa_elig_summ_eligibility_result)
			msa_elig_summ_reporting_status = trim(msa_elig_summ_reporting_status)
			msa_elig_summ_benefit = trim(msa_elig_summ_benefit)
			msa_elig_summ_shared_houshold = trim(msa_elig_summ_shared_houshold)

			If msa_elig_summ_vendor_reason_code = "01" Then msa_elig_summ_vendor_reason_info = "Client Request"
			If msa_elig_summ_vendor_reason_code = "05" Then msa_elig_summ_vendor_reason_info = "Money Mismanagement"
			If msa_elig_summ_vendor_reason_code = "09" Then msa_elig_summ_vendor_reason_info = "Emergency"
			If msa_elig_summ_vendor_reason_code = "10" Then msa_elig_summ_vendor_reason_info = "Chemical Dependency"
			If msa_elig_summ_vendor_reason_code = "11" Then msa_elig_summ_vendor_reason_info = "No Residence"
			If msa_elig_summ_vendor_reason_code = "20" Then msa_elig_summ_vendor_reason_info = "Grant Diversion"

			msa_elig_summ_total_assets = trim(msa_elig_summ_total_assets)
			msa_elig_summ_maximum_assets = trim(msa_elig_summ_maximum_assets)
			msa_elig_summ_grant = trim(msa_elig_summ_grant)
			msa_elig_summ_current_payment = trim(msa_elig_summ_current_payment)

			msa_elig_summ_worker_message = trim(msa_elig_summ_worker_message)
		End if

		Call back_to_SELF
	end sub
end class

class ga_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found
	public er_month
	public hrf_month
	public er_status
	public er_caf_date
	public er_interview_date
	public hrf_status
	public hrf_doc_date

	public ga_elig_case_status
	public ga_elig_file_unit_type_code
	public ga_elig_faci_file_unit_type_code
	public ga_elig_file_unit_type_info
	public ga_elig_faci_file_unit_type_info

	public ga_elig_ref_numbs()
	public ga_elig_membs_full_name()
	public ga_elig_membs_relationship_code()
	public ga_elig_membs_relationship_info()
	public ga_elig_membs_code()
	public ga_elig_membs_info()
	public ga_elig_membs_elig_basis_code()
	public ga_elig_membs_elig_basis_info()
	public ga_elig_membs_elig_status()
	public ga_elig_membs_budget_cycle()
	public ga_elig_membs_elig_begin_date()
	public ga_elig_membs_test_absence()
	public ga_elig_membs_test_dupl_assistance()
	public ga_elig_membs_test_ga_coop()
	public ga_elig_membs_test_ive()
	public ga_elig_membs_test_ssi()
	public ga_elig_membs_test_lump_sum_payment()
	public ga_elig_membs_test_unit_member()
	public ga_elig_membs_test_imig_status_verif()
	public ga_elig_membs_test_imig_status()
	public ga_elig_membs_test_basis_of_elig()
	public ga_elig_membs_test_elig_other_prgm()
	public ga_elig_membs_test_ssn_coop()

	public ga_elig_case_test_appl_withdrawn
	public ga_elig_case_test_dupl_assistance
	public ga_elig_case_test_fail_coop
	public ga_elig_case_test_fail_file
	public ga_elig_case_test_eligible_member
	public ga_elig_case_test_prosp_net_income
	public ga_elig_case_test_retro_net_income
	public ga_elig_case_test_residence
	public ga_elig_case_test_assets
	public ga_elig_case_test_eligible_other_prgm
	public ga_elig_case_test_verif
	public ga_elig_case_test_lump_sum_payment

	public ga_elig_case_budg_gross_wages
	public ga_elig_case_budg_gross_self_emp
	public ga_elig_case_budg_total_gross_income
	public ga_elig_case_budg_standard_EI_disregard
	public ga_elig_case_budg_earned_income_subtotal
	public ga_elig_case_budg_earned_income_disregard_percent
	public ga_elig_case_budg_earned_income_disregard_amount
	public ga_elig_case_budg_total_deductions
	public ga_elig_case_budg_net_earned_income
	public ga_elig_case_budg_unearned_income
	public ga_elig_case_budg_counted_school_income
	public ga_elig_case_budg_total_deemed_income
	public ga_elig_case_budg_total_countable_income

	public ga_elig_case_budg_payment_standard
	public ga_elig_case_budg_payment_subtotal
	public ga_elig_case_budg_prorated_from
	public ga_elig_case_budg_prorated_to
	public ga_elig_case_budg_grant_subtotal
	public ga_elig_case_budg_total_assets
	public ga_elig_case_budg_ga_exclusion
	public ga_elig_case_budg_countable_assets
	public ga_elig_case_budg_maximum_assets
	public ga_elig_case_budg_reason_ga_exclusion
	public ga_elig_case_budg_pers_needs_payment_standard
	public ga_elig_case_budg_pers_needs_payment_subtotal
	public ga_elig_case_budg_pers_needs_prorated_from
	public ga_elig_case_budg_pers_needs_prorated_to
	public ga_elig_case_budg_pers_needs_grant_subtotal
	public ga_elig_case_budg_total_ga_grant_amount

	public ga_elig_summ_approved_date
	public ga_elig_summ_process_date
	public ga_elig_summ_date_last_approval
	public ga_elig_summ_curr_prog_status
	public ga_elig_summ_eligibility_result
	public ga_elig_summ_hrf_reporting
	public ga_elig_summ_source_of_info
	public ga_elig_summ_eligibility_begin_date
	public ga_elig_summ_eligiblity_review_date
	public ga_elig_summ_budget_cycle
	public ga_elig_summ_filing_unit_type_code
	public ga_elig_summ_filing_unit_type_info
	public ga_elig_summ_faci_unit_type_code
	public ga_elig_summ_faci_unit_type_info
	public ga_elig_summ_responsible_county
	public ga_elig_summ_vendor_reason_code
	public ga_elig_summ_vendor_reason_info
	public ga_elig_summ_total_assets
	public ga_elig_summ_client_faci_obligation
	public ga_elig_summ_standards
	public ga_elig_summ_counted_income
	public ga_elig_summ_monthly_grant
	public ga_elig_summ_amount_to_be_paid
	public ga_elig_summ_action_code
	public ga_elig_summ_action_info
	public ga_elig_summ_reason_code
	public ga_elig_summ_reason_info
	public ga_elig_summ_worker_message

	public sub read_elig()
		approved_today = False
		approved_version_found = False

		call navigate_to_MAXIS_screen("ELIG", "GA  ")
		EMWriteScreen elig_footer_month, 20, 54
		EMWriteScreen elig_footer_year, 20, 57
		Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

	 		EMReadScreen ga_elig_case_status, 12, 18, 23
			EMReadScreen ga_elig_file_unit_type_code, 1, 18, 52
			EMReadScreen ga_elig_faci_file_unit_type_code, 1, 18, 77

			ga_elig_case_status = trim(ga_elig_case_status)

			If ga_elig_file_unit_type_code = "1" Then ga_elig_file_unit_type_info = "Single Adult"
			If ga_elig_file_unit_type_code = "2" Then ga_elig_file_unit_type_info = "Single Adult living with Parents"
			If ga_elig_file_unit_type_code = "3" Then ga_elig_file_unit_type_info = "Minor Child Outside the Home"
			If ga_elig_file_unit_type_code = "6" Then ga_elig_file_unit_type_info = "Married Couple"
			If ga_elig_file_unit_type_code = "9" Then ga_elig_file_unit_type_info = "Family State Food Program"

			If ga_elig_faci_file_unit_type_code = "5" Then ga_elig_faci_file_unit_type_info = "Personal Needs"


			ReDim ga_elig_ref_numbs(0)
			ReDim ga_elig_membs_full_name(0)
			ReDim ga_elig_membs_relationship_code(0)
			ReDim ga_elig_membs_relationship_info(0)
			ReDim ga_elig_membs_code(0)
			ReDim ga_elig_membs_info(0)
			ReDim ga_elig_membs_elig_basis_code(0)
			ReDim ga_elig_membs_elig_basis_info(0)
			ReDim ga_elig_membs_elig_status(0)
			ReDim ga_elig_membs_budget_cycle(0)
			ReDim ga_elig_membs_elig_begin_date(0)
			ReDim ga_elig_membs_test_absence(0)
			ReDim ga_elig_membs_test_dupl_assistance(0)
			ReDim ga_elig_membs_test_ga_coop(0)
			ReDim ga_elig_membs_test_ive(0)
			ReDim ga_elig_membs_test_ssi(0)
			ReDim ga_elig_membs_test_lump_sum_payment(0)
			ReDim ga_elig_membs_test_unit_member(0)
			ReDim ga_elig_membs_test_imig_status_verif(0)
			ReDim ga_elig_membs_test_imig_status(0)
			ReDim ga_elig_membs_test_basis_of_elig(0)
			ReDim ga_elig_membs_test_elig_other_prgm(0)
			ReDim ga_elig_membs_test_ssn_coop(0)

			elig_memb_count = 0
			ga_row = 8
			Do
				EMReadScreen ref_numb, 2, ga_row, 9

				ReDim preserve ga_elig_ref_numbs(elig_memb_count)
				ReDim preserve ga_elig_membs_full_name(elig_memb_count)
				ReDim preserve ga_elig_membs_relationship_code(elig_memb_count)
				ReDim preserve ga_elig_membs_relationship_info(elig_memb_count)
				ReDim preserve ga_elig_membs_code(elig_memb_count)
				ReDim preserve ga_elig_membs_info(elig_memb_count)
				ReDim preserve ga_elig_membs_elig_basis_code(elig_memb_count)
				ReDim preserve ga_elig_membs_elig_basis_info(elig_memb_count)
				ReDim preserve ga_elig_membs_elig_status(elig_memb_count)
				ReDim preserve ga_elig_membs_budget_cycle(elig_memb_count)
				ReDim preserve ga_elig_membs_elig_begin_date(elig_memb_count)
				ReDim preserve ga_elig_membs_test_absence(elig_memb_count)
				ReDim preserve ga_elig_membs_test_dupl_assistance(elig_memb_count)
				ReDim preserve ga_elig_membs_test_ga_coop(elig_memb_count)
				ReDim preserve ga_elig_membs_test_ive(elig_memb_count)
				ReDim preserve ga_elig_membs_test_ssi(elig_memb_count)
				ReDim preserve ga_elig_membs_test_lump_sum_payment(elig_memb_count)
				ReDim preserve ga_elig_membs_test_unit_member(elig_memb_count)
				ReDim preserve ga_elig_membs_test_imig_status_verif(elig_memb_count)
				ReDim preserve ga_elig_membs_test_imig_status(elig_memb_count)
				ReDim preserve ga_elig_membs_test_basis_of_elig(elig_memb_count)
				ReDim preserve ga_elig_membs_test_elig_other_prgm(elig_memb_count)
				ReDim preserve ga_elig_membs_test_ssn_coop(elig_memb_count)

				ga_elig_ref_numbs(elig_memb_count) = ref_numb
				EMReadScreen full_name_information, 20, ga_row, 12
				full_name_information = trim(full_name_information)
				name_array = split(full_name_information, " ")
				For each name_parts in name_array
					If len(name_parts) <> 1 Then ga_elig_membs_full_name(elig_memb_count) = ga_elig_membs_full_name(elig_memb_count) & " " & name_parts
				Next
				ga_elig_membs_full_name(elig_memb_count) = trim((ga_elig_membs_full_name(elig_memb_count)))
				EMReadScreen ga_elig_membs_relationship_code(elig_memb_count), 2, ga_row, 33


				If ga_elig_membs_relationship_code(elig_memb_count) = "01" Then ga_elig_membs_relationship_info(elig_memb_count) = "Applicant"
				If ga_elig_membs_relationship_code(elig_memb_count) = "02" Then ga_elig_membs_relationship_info(elig_memb_count) = "Spouse"
				If ga_elig_membs_relationship_code(elig_memb_count) = "03" Then ga_elig_membs_relationship_info(elig_memb_count) = "Child"
				If ga_elig_membs_relationship_code(elig_memb_count) = "04" Then ga_elig_membs_relationship_info(elig_memb_count) = "Parent"
				If ga_elig_membs_relationship_code(elig_memb_count) = "05" Then ga_elig_membs_relationship_info(elig_memb_count) = "Sibling"
				If ga_elig_membs_relationship_code(elig_memb_count) = "06" Then ga_elig_membs_relationship_info(elig_memb_count) = "Step Sibling"
				If ga_elig_membs_relationship_code(elig_memb_count) = "08" Then ga_elig_membs_relationship_info(elig_memb_count) = "Step Child"
				If ga_elig_membs_relationship_code(elig_memb_count) = "09" Then ga_elig_membs_relationship_info(elig_memb_count) = "Step Parent"
				If ga_elig_membs_relationship_code(elig_memb_count) = "10" Then ga_elig_membs_relationship_info(elig_memb_count) = "Aunt"
				If ga_elig_membs_relationship_code(elig_memb_count) = "11" Then ga_elig_membs_relationship_info(elig_memb_count) = "Uncle"
				If ga_elig_membs_relationship_code(elig_memb_count) = "12" Then ga_elig_membs_relationship_info(elig_memb_count) = "Niece"
				If ga_elig_membs_relationship_code(elig_memb_count) = "13" Then ga_elig_membs_relationship_info(elig_memb_count) = "Nephew"
				If ga_elig_membs_relationship_code(elig_memb_count) = "14" Then ga_elig_membs_relationship_info(elig_memb_count) = "Cousin"
				If ga_elig_membs_relationship_code(elig_memb_count) = "15" Then ga_elig_membs_relationship_info(elig_memb_count) = "Grandparent"
				If ga_elig_membs_relationship_code(elig_memb_count) = "16" Then ga_elig_membs_relationship_info(elig_memb_count) = "Grandchild"
				If ga_elig_membs_relationship_code(elig_memb_count) = "17" Then ga_elig_membs_relationship_info(elig_memb_count) = "Other Relative"
				If ga_elig_membs_relationship_code(elig_memb_count) = "18" Then ga_elig_membs_relationship_info(elig_memb_count) = "Legal Guardian"
				If ga_elig_membs_relationship_code(elig_memb_count) = "24" Then ga_elig_membs_relationship_info(elig_memb_count) = "Not Related"
				If ga_elig_membs_relationship_code(elig_memb_count) = "25" Then ga_elig_membs_relationship_info(elig_memb_count) = "Live-In Attendant"
				If ga_elig_membs_relationship_code(elig_memb_count) = "27" Then ga_elig_membs_relationship_info(elig_memb_count) = "Unknown"

				EMReadScreen ga_elig_membs_code(elig_memb_count), 1, ga_row, 48

				If ga_elig_membs_code(elig_memb_count) = "A" Then ga_elig_membs_info(elig_memb_count) = "Assistance Unit Member"
				If ga_elig_membs_code(elig_memb_count) = "C" Then ga_elig_membs_info(elig_memb_count) = "Deemer"
				If ga_elig_membs_code(elig_memb_count) = "F" Then ga_elig_membs_info(elig_memb_count) = "Ineligible - Counted without Deductions"
				If ga_elig_membs_code(elig_memb_count) = "S" Then ga_elig_membs_info(elig_memb_count) = "Ineligible - Counted with Deduction"
				If ga_elig_membs_code(elig_memb_count) = "G" Then ga_elig_membs_info(elig_memb_count) = "Ineligible Affects Grant"
				If ga_elig_membs_code(elig_memb_count) = "I" Then ga_elig_membs_info(elig_memb_count) = "Ineligible Par of Unit"
				If ga_elig_membs_code(elig_memb_count) = "L" Then ga_elig_membs_info(elig_memb_count) = "Other Adult Applicant"
				If ga_elig_membs_code(elig_memb_count) = "M" Then ga_elig_membs_info(elig_memb_count) = "Allocation Only"
				If ga_elig_membs_code(elig_memb_count) = "N" Then ga_elig_membs_info(elig_memb_count) = "Not Counted"
				If ga_elig_membs_code(elig_memb_count) = "U" Then ga_elig_membs_info(elig_memb_count) = "Unknown"

				EMReadScreen ga_elig_membs_elig_basis_code(elig_memb_count), 2, row, 52

				If ga_elig_membs_elig_basis_code(elig_memb_count) = "04" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Permanent Ill Or Incap"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "05" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Temporary Ill Or Incap"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "06" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Care Of Ill Or Incap Mbr"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "07" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Requires Services In Residence"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "09" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Mntl Ill Or Dev Disabled"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "10" then ga_elig_membs_elig_basis_info(elig_memb_count) = "SSI/RSDI Pend"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "11" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Appealing SSI/RSDI Denial"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "12" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Advanced Age"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "13" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Learning Disability"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "17" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Protect/Court Ordered"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "20" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Age 16 Or 17 SS Approval"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "25" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Emancipated Minor"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "28" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Unemployable"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "29" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Displaced Hmkr(Ft Student)"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "30" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Minor W/ Adult Unrelated"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "32" then ga_elig_membs_elig_basis_info(elig_memb_count) = "ESL, Adult/HS At Least Half Time, Adult"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "35" then ga_elig_membs_elig_basis_info(elig_memb_count) = "Drug/Alcohol Addiction(DAA)"
				If ga_elig_membs_elig_basis_code(elig_memb_count) = "99" then ga_elig_membs_elig_basis_info(elig_memb_count) = "No Elig Basis"

				EMReadScreen ga_elig_membs_elig_status(elig_memb_count), 4, row, 57

				If ga_elig_membs_elig_status(elig_memb_count) = "ELIG" then ga_elig_membs_elig_status(elig_memb_count) = "ELIGIBLE"
				If ga_elig_membs_elig_status(elig_memb_count) = "INEL" then ga_elig_membs_elig_status(elig_memb_count) = "INELIGIBLE"

				EMReadScreen ga_elig_membs_budget_cycle(elig_memb_count), 1, row, 63

				If ga_elig_membs_budget_cycle(elig_memb_count) = "P" then ga_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
				If ga_elig_membs_budget_cycle(elig_memb_count) = "R" then ga_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"

				EMReadScreen ga_elig_membs_elig_begin_date(elig_memb_count), 8, row, 66

				Call write_value_and_transmit("X", row, 6)

				EMReadScreen ga_elig_membs_test_absence(elig_memb_count), 			6, 11, 12
				EMReadScreen ga_elig_membs_test_dupl_assistance(elig_memb_count), 	6, 12, 12
				EMReadScreen ga_elig_membs_test_ga_coop(elig_memb_count), 			6, 13, 12
				EMReadScreen ga_elig_membs_test_ive(elig_memb_count), 				6, 14, 12
				EMReadScreen ga_elig_membs_test_ssi(elig_memb_count), 				6, 15, 12
				EMReadScreen ga_elig_membs_test_lump_sum_payment(elig_memb_count), 	6, 16, 12


				EMReadScreen ga_elig_membs_test_unit_member(elig_memb_count), 		6, 11, 42
				EMReadScreen ga_elig_membs_test_imig_status_verif(elig_memb_count), 6, 12, 42
				EMReadScreen ga_elig_membs_test_imig_status(elig_memb_count), 		6, 13, 42
				EMReadScreen ga_elig_membs_test_basis_of_elig(elig_memb_count), 	6, 14, 42
				EMReadScreen ga_elig_membs_test_elig_other_prgm(elig_memb_count), 	6, 15, 42
				EMReadScreen ga_elig_membs_test_ssn_coop(elig_memb_count), 			6, 16, 42

				ga_elig_membs_test_absence(elig_memb_count) = trim(ga_elig_membs_test_absence(elig_memb_count))
				ga_elig_membs_test_dupl_assistance(elig_memb_count) = trim(ga_elig_membs_test_dupl_assistance(elig_memb_count))
				ga_elig_membs_test_ga_coop(elig_memb_count) = trim(ga_elig_membs_test_ga_coop(elig_memb_count))
				ga_elig_membs_test_ive(elig_memb_count) = trim(ga_elig_membs_test_ive(elig_memb_count))
				ga_elig_membs_test_ssi(elig_memb_count) = trim(ga_elig_membs_test_ssi(elig_memb_count))
				ga_elig_membs_test_lump_sum_payment(elig_memb_count) = trim(ga_elig_membs_test_lump_sum_payment(elig_memb_count))

				ga_elig_membs_test_unit_member(elig_memb_count) = trim(ga_elig_membs_test_unit_member(elig_memb_count))
				ga_elig_membs_test_imig_status_verif(elig_memb_count) = trim(ga_elig_membs_test_imig_status_verif(elig_memb_count))
				ga_elig_membs_test_imig_status(elig_memb_count) = trim(ga_elig_membs_test_imig_status(elig_memb_count))
				ga_elig_membs_test_basis_of_elig(elig_memb_count) = trim(ga_elig_membs_test_basis_of_elig(elig_memb_count))
				ga_elig_membs_test_elig_other_prgm(elig_memb_count) = trim(ga_elig_membs_test_elig_other_prgm(elig_memb_count))
				ga_elig_membs_test_ssn_coop(elig_memb_count) = trim(ga_elig_membs_test_ssn_coop(elig_memb_count))

				transmit

				ga_row = ga_row + 1
				elig_memb_count = elig_memb_count + 1
				EMReadScreen next_ref_numb, 2, ga_row, 9
			Loop until next_ref_numb = "  "

			transmit 		'going to the next panel - GACR

			EMReadScreen ga_elig_case_test_appl_withdrawn, 		6, 8, 10
			EMReadScreen ga_elig_case_test_dupl_assistance, 	6, 9, 10
			EMReadScreen ga_elig_case_test_fail_coop, 			6, 10, 10
			EMReadScreen ga_elig_case_test_fail_file, 			6, 11, 10
			EMReadScreen ga_elig_case_test_eligible_member, 	6, 12, 10
			EMReadScreen ga_elig_case_test_prosp_net_income, 	6, 13, 10

			EMReadScreen ga_elig_case_test_retro_net_income, 	6, 8, 46
			EMReadScreen ga_elig_case_test_residence, 			6, 9, 46
			EMReadScreen ga_elig_case_test_assets, 				6, 10, 46
			EMReadScreen ga_elig_case_test_eligible_other_prgm, 6, 11, 46
			EMReadScreen ga_elig_case_test_verif, 				6, 12, 46
			EMReadScreen ga_elig_case_test_lump_sum_payment, 	6, 13, 46

			ga_elig_case_test_appl_withdrawn = trim(ga_elig_case_test_appl_withdrawn)
			ga_elig_case_test_dupl_assistance = trim(ga_elig_case_test_dupl_assistance)
			ga_elig_case_test_fail_coop = trim(ga_elig_case_test_fail_coop)
			ga_elig_case_test_fail_file = trim(ga_elig_case_test_fail_file)
			ga_elig_case_test_eligible_member = trim(ga_elig_case_test_eligible_member)
			ga_elig_case_test_prosp_net_income = trim(ga_elig_case_test_prosp_net_income)

			ga_elig_case_test_retro_net_income = trim(ga_elig_case_test_retro_net_income)
			ga_elig_case_test_residence = trim(ga_elig_case_test_residence)
			ga_elig_case_test_assets = trim(ga_elig_case_test_assets)
			ga_elig_case_test_eligible_other_prgm = trim(ga_elig_case_test_eligible_other_prgm)
			ga_elig_case_test_verif = trim(ga_elig_case_test_verif)
			ga_elig_case_test_lump_sum_payment = trim(ga_elig_case_test_lump_sum_payment)

			' Call write_value_and_transmit("X", 13, 4)		'This is the Prosp Net Income Pop-Up - this appears to match the information on GAb1 - so we are not reading it'

			transmit 		'going to the next panel - GAB1

			EMReadScreen ga_elig_case_budg_gross_wages, 					10, 6, 29
			EMReadScreen ga_elig_case_budg_gross_self_emp, 					10, 7, 29
			EMReadScreen ga_elig_case_budg_total_gross_income, 				10, 9, 29
			EMReadScreen ga_elig_case_budg_standard_EI_disregard, 			10, 13, 29
			EMReadScreen ga_elig_case_budg_earned_income_subtotal, 			10, 14, 29
			EMReadScreen ga_elig_case_budg_earned_income_disregard_percent, 2, 15, 23
			EMReadScreen ga_elig_case_budg_earned_income_disregard_amount, 	10, 15, 29
			EMReadScreen ga_elig_case_budg_total_deductions, 				10, 17, 29

			EMReadScreen ga_elig_case_budg_net_earned_income, 				10, 6, 71
			EMReadScreen ga_elig_case_budg_unearned_income, 				10, 8, 71
			EMReadScreen ga_elig_case_budg_counted_school_income, 			10, 10, 71
			EMReadScreen ga_elig_case_budg_total_deemed_income, 			10, 14, 71
			EMReadScreen ga_elig_case_budg_total_countable_income, 			10, 17, 71

			ga_elig_case_budg_gross_wages = trim(ga_elig_case_budg_gross_wages)
			ga_elig_case_budg_gross_self_emp = trim(ga_elig_case_budg_gross_self_emp)
			ga_elig_case_budg_total_gross_income = trim(ga_elig_case_budg_total_gross_income)
			ga_elig_case_budg_standard_EI_disregard = trim(ga_elig_case_budg_standard_EI_disregard)
			ga_elig_case_budg_earned_income_subtotal = trim(ga_elig_case_budg_earned_income_subtotal)
			ga_elig_case_budg_earned_income_disregard_percent = trim(ga_elig_case_budg_earned_income_disregard_percent)
			ga_elig_case_budg_earned_income_disregard_amount = trim(ga_elig_case_budg_earned_income_disregard_amount)
			ga_elig_case_budg_total_deductions = trim(ga_elig_case_budg_total_deductions)

			ga_elig_case_budg_net_earned_income = trim(ga_elig_case_budg_net_earned_income)
			ga_elig_case_budg_unearned_income = trim(ga_elig_case_budg_unearned_income)
			ga_elig_case_budg_counted_school_income = trim(ga_elig_case_budg_counted_school_income)
			ga_elig_case_budg_total_deemed_income = trim(ga_elig_case_budg_total_deemed_income)
			ga_elig_case_budg_total_countable_income = trim(ga_elig_case_budg_total_countable_income)

			transmit 		'going to the next panel - GAB2

			EMReadScreen ga_elig_case_budg_payment_standard, 	10, 6, 34
			' EMReadScreen ga_elig_case_budg_total_countable_income, 10, 7, 34
			EMReadScreen ga_elig_case_budg_payment_subtotal, 	10, 8, 34
			EMReadScreen ga_elig_case_budg_prorated_from, 		5, 10, 15
			EMReadScreen ga_elig_case_budg_prorated_to, 		5, 10, 25
			EMReadScreen ga_elig_case_budg_grant_subtotal, 		10, 11, 34
			EMReadScreen ga_elig_case_budg_total_assets, 		10, 14, 34
			EMReadScreen ga_elig_case_budg_ga_exclusion, 		10, 15, 34
			EMReadScreen ga_elig_case_budg_countable_assets, 	10, 16, 34
			EMReadScreen ga_elig_case_budg_maximum_assets, 		10, 17, 34
			EMReadScreen ga_elig_case_budg_reason_ga_exclusion, 10, 18, 34

			EMReadScreen ga_elig_case_budg_pers_needs_payment_standard, 10, 6, 72
			' EMReadScreen ga_elig_case_budg_total_countable_income, 10, 7, 72
			EMReadScreen ga_elig_case_budg_pers_needs_payment_subtotal, 10, 8, 72
			EMReadScreen ga_elig_case_budg_pers_needs_prorated_from, 	5, 10, 58
			EMReadScreen ga_elig_case_budg_pers_needs_prorated_to, 		5, 10, 68
			EMReadScreen ga_elig_case_budg_pers_needs_grant_subtotal, 	10, 11, 72
			EMReadScreen ga_elig_case_budg_total_ga_grant_amount, 		10, 13, 72

			ga_elig_case_budg_payment_standard = trim(ga_elig_case_budg_payment_standard)
			ga_elig_case_budg_payment_subtotal = trim(ga_elig_case_budg_payment_subtotal)
			ga_elig_case_budg_prorated_from = trim(ga_elig_case_budg_prorated_from)
			ga_elig_case_budg_prorated_to = trim(ga_elig_case_budg_prorated_to)
			ga_elig_case_budg_grant_subtotal = trim(ga_elig_case_budg_grant_subtotal)
			ga_elig_case_budg_total_assets = trim(ga_elig_case_budg_total_assets)
			ga_elig_case_budg_ga_exclusion = trim(ga_elig_case_budg_ga_exclusion)
			ga_elig_case_budg_countable_assets = trim(ga_elig_case_budg_countable_assets)
			ga_elig_case_budg_maximum_assets = trim(ga_elig_case_budg_maximum_assets)
			ga_elig_case_budg_reason_ga_exclusion = trim(ga_elig_case_budg_reason_ga_exclusion)

			ga_elig_case_budg_pers_needs_payment_standard = trim(ga_elig_case_budg_pers_needs_payment_standard)
			ga_elig_case_budg_pers_needs_payment_subtotal = trim(ga_elig_case_budg_pers_needs_payment_subtotal)
			ga_elig_case_budg_pers_needs_prorated_from = trim(ga_elig_case_budg_pers_needs_prorated_from)
			ga_elig_case_budg_pers_needs_prorated_to = trim(ga_elig_case_budg_pers_needs_prorated_to)
			ga_elig_case_budg_pers_needs_grant_subtotal = trim(ga_elig_case_budg_pers_needs_grant_subtotal)
			ga_elig_case_budg_total_ga_grant_amount = trim(ga_elig_case_budg_total_ga_grant_amount)

			If ga_elig_case_budg_prorated_from <> "" Then
				ga_elig_case_budg_prorated_from = replace(ga_elig_case_budg_prorated_from, " ", "/")
				ga_elig_case_budg_prorated_from = ga_elig_case_budg_prorated_from & "/" & elig_footer_year
			End If
			If ga_elig_case_budg_prorated_to <> "" Then
				ga_elig_case_budg_prorated_to = replace(ga_elig_case_budg_prorated_to, " ", "/")
				ga_elig_case_budg_prorated_to = ga_elig_case_budg_prorated_to & "/" & elig_footer_year
			End If
			If ga_elig_case_budg_pers_needs_prorated_from <> "" Then
				ga_elig_case_budg_pers_needs_prorated_from = replace(ga_elig_case_budg_pers_needs_prorated_from, " ", "/")
				ga_elig_case_budg_pers_needs_prorated_from = ga_elig_case_budg_pers_needs_prorated_from & "/" & elig_footer_year
			End If
			If ga_elig_case_budg_pers_needs_prorated_to <> "" Then
				ga_elig_case_budg_pers_needs_prorated_to = replace(ga_elig_case_budg_pers_needs_prorated_to, " ", "/")
				ga_elig_case_budg_pers_needs_prorated_to = ga_elig_case_budg_pers_needs_prorated_to & "/" & elig_footer_year
			End If

			transmit 		'going to the next panel - GASM

			EMReadScreen ga_elig_summ_approved_date, 8, 3, 15
			EMReadScreen ga_elig_summ_process_date, 8, 2, 73
			EMReadScreen ga_elig_summ_date_last_approval, 8, 5, 32
			EMReadScreen ga_elig_summ_curr_prog_status, 12, 6, 32
			EMReadScreen ga_elig_summ_eligibility_result, 12, 7, 32
			EMReadScreen ga_elig_summ_hrf_reporting, 12, 8, 32
			EMReadScreen ga_elig_summ_source_of_info, 4, 9, 32
			EMReadScreen ga_elig_summ_eligibility_begin_date, 8, 10, 32
			EMReadScreen ga_elig_summ_eligiblity_review_date, 8, 11, 32
			EMReadScreen ga_elig_summ_budget_cycle, 5, 12, 32
			EMReadScreen ga_elig_summ_filing_unit_type_code, 1, 13, 32
			EMReadScreen ga_elig_summ_faci_unit_type_code, 1, 14, 32
			EMReadScreen ga_elig_summ_responsible_county, 2, 15, 32
			EMReadScreen ga_elig_summ_vendor_reason_code, 2, 16, 32

			EMReadScreen ga_elig_summ_total_assets, 10, 5, 71
			EMReadScreen ga_elig_summ_client_faci_obligation, 10, 6, 71
			EMReadScreen ga_elig_summ_standards, 10, 7, 71
			EMReadScreen ga_elig_summ_counted_income, 10, 8, 71
			EMReadScreen ga_elig_summ_monthly_grant, 10, 9, 71
			EMReadScreen ga_elig_summ_amount_to_be_paid, 10, 14, 71
			EMReadScreen ga_elig_summ_action_code, 1, 15, 53
			EMReadScreen ga_elig_summ_reason_code, 2, 16, 53

			EMReadScreen ga_elig_summ_worker_message, 80, 19, 1

			ga_elig_summ_curr_prog_status = trim(ga_elig_summ_curr_prog_status)
			ga_elig_summ_eligibility_result = trim(ga_elig_summ_eligibility_result)
			ga_elig_summ_hrf_reporting = trim(ga_elig_summ_hrf_reporting)

			If ga_elig_summ_filing_unit_type_code = "1" Then ga_elig_summ_filing_unit_type_info = "Single Adult"
			If ga_elig_summ_filing_unit_type_code = "2" Then ga_elig_summ_filing_unit_type_info = "Single Adult Lv W/ Parents"
			If ga_elig_summ_filing_unit_type_code = "3" Then ga_elig_summ_filing_unit_type_info = "Minor Child Outside Home"
			If ga_elig_summ_filing_unit_type_code = "6" Then ga_elig_summ_filing_unit_type_info = "Married Couple"
			If ga_elig_summ_filing_unit_type_code = "9" Then ga_elig_summ_filing_unit_type_info = "Family State Food Program"

			If ga_elig_summ_faci_unit_type_code = "5" Then ga_elig_summ_faci_unit_type_info = "Personal Needs"

			If ga_elig_summ_vendor_reason_code = "01" Then ga_elig_summ_vendor_reason_info = "Client Request"
			If ga_elig_summ_vendor_reason_code = "05" Then ga_elig_summ_vendor_reason_info = "Money Mismanagement"
			If ga_elig_summ_vendor_reason_code = "09" Then ga_elig_summ_vendor_reason_info = "Emergency"
			If ga_elig_summ_vendor_reason_code = "10" Then ga_elig_summ_vendor_reason_info = "Chemical Dependency"
			If ga_elig_summ_vendor_reason_code = "11" Then ga_elig_summ_vendor_reason_info = "No Residence"
			If ga_elig_summ_vendor_reason_code = "20" Then ga_elig_summ_vendor_reason_info = "Grant Diversion"


			ga_elig_summ_total_assets = trim(ga_elig_summ_total_assets)
			ga_elig_summ_client_faci_obligation = trim(ga_elig_summ_client_faci_obligation)
			ga_elig_summ_standards = trim(ga_elig_summ_standards)
			ga_elig_summ_counted_income = trim(ga_elig_summ_counted_income)
			ga_elig_summ_monthly_grant = trim(ga_elig_summ_monthly_grant)
			ga_elig_summ_amount_to_be_paid = trim(ga_elig_summ_amount_to_be_paid)

			If ga_elig_summ_action_code = "1" Then ga_elig_summ_action_info = "Open"
			If ga_elig_summ_action_code = "2" Then ga_elig_summ_action_info = "Suspend"
			If ga_elig_summ_action_code = "3" Then ga_elig_summ_action_info = "Unsuspend"
			If ga_elig_summ_action_code = "4" Then ga_elig_summ_action_info = "Review - Grant Change"
			If ga_elig_summ_action_code = "5" Then ga_elig_summ_action_info = "Close"
			If ga_elig_summ_action_code = "7" Then ga_elig_summ_action_info = "Grant Change - Chng Reported"
			If ga_elig_summ_action_code = "8" Then ga_elig_summ_action_info = "Review - No Grant Chng"
			If ga_elig_summ_action_code = "9" Then ga_elig_summ_action_info = "No Grant Chng - Chng Reported"
			If ga_elig_summ_action_code = "0" Then ga_elig_summ_action_info = "STAT Change - No Notice Rqrd"
			If ga_elig_summ_action_code = "C" Then ga_elig_summ_action_info = "Reinstate Closed Case"

			If ga_elig_summ_reason_code = "01" Then ga_elig_summ_reason_info = "Earned Income Increased"
			If ga_elig_summ_reason_code = "02" Then ga_elig_summ_reason_info = "Earned Income Decreased"
			If ga_elig_summ_reason_code = "03" Then ga_elig_summ_reason_info = "Unearned Income Increased"
			If ga_elig_summ_reason_code = "04" Then ga_elig_summ_reason_info = "Unearned Income Decreased"
			If ga_elig_summ_reason_code = "05" Then ga_elig_summ_reason_info = "Expenses/Deductions Increased"
			If ga_elig_summ_reason_code = "06" Then ga_elig_summ_reason_info = "Expenses/Deductions Decr"
			If ga_elig_summ_reason_code = "08" Then ga_elig_summ_reason_info = "No Proof Given"
			If ga_elig_summ_reason_code = "09" Then ga_elig_summ_reason_info = "Did Not Return Review Form"
			If ga_elig_summ_reason_code = "10" Then ga_elig_summ_reason_info = "Non Coop With GA Rules"
			If ga_elig_summ_reason_code = "12" Then ga_elig_summ_reason_info = "Must Apply For Other Benefit"
			If ga_elig_summ_reason_code = "14" Then ga_elig_summ_reason_info = "Not At Given Address"
			If ga_elig_summ_reason_code = "16" Then ga_elig_summ_reason_info = "Request Close"
			If ga_elig_summ_reason_code = "17" Then ga_elig_summ_reason_info = "Eligibility For Other Cash Program"
			If ga_elig_summ_reason_code = "18" Then ga_elig_summ_reason_info = "Non State Resident"
			If ga_elig_summ_reason_code = "19" Then ga_elig_summ_reason_info = "Client Died"
			If ga_elig_summ_reason_code = "20" Then ga_elig_summ_reason_info = "Household Member Died"
			If ga_elig_summ_reason_code = "22" Then ga_elig_summ_reason_info = "Excess Income"
			If ga_elig_summ_reason_code = "23" Then ga_elig_summ_reason_info = "Assets over the GA Limit"
			If ga_elig_summ_reason_code = "24" Then ga_elig_summ_reason_info = "Tranfer of Assets - No GA Eligiblity"
			If ga_elig_summ_reason_code = "27" Then ga_elig_summ_reason_info = "Fail To Sign Interim Assistance Agreemnt"
			If ga_elig_summ_reason_code = "28" Then ga_elig_summ_reason_info = "Program Reqquirements Have Been Met"
			If ga_elig_summ_reason_code = "30" Then ga_elig_summ_reason_info = "Household Size Change"
			If ga_elig_summ_reason_code = "31" Then ga_elig_summ_reason_info = "Review - No Change"
			If ga_elig_summ_reason_code = "32" Then ga_elig_summ_reason_info = "Begin Recoupment"
			If ga_elig_summ_reason_code = "33" Then ga_elig_summ_reason_info = "Change Recoupment"
			If ga_elig_summ_reason_code = "34" Then ga_elig_summ_reason_info = "End Recoupment"
			If ga_elig_summ_reason_code = "35" Then ga_elig_summ_reason_info = "New GA Basis Of Eligiblity"
			If ga_elig_summ_reason_code = "36" Then ga_elig_summ_reason_info = "Add/Change/Delete Vendor"
			If ga_elig_summ_reason_code = "39" Then ga_elig_summ_reason_info = "Person In/Out Facility"
			If ga_elig_summ_reason_code = "49" Then ga_elig_summ_reason_info = "No HRF"
			If ga_elig_summ_reason_code = "51" Then ga_elig_summ_reason_info = "Under Control Of Penal System"
			If ga_elig_summ_reason_code = "52" Then ga_elig_summ_reason_info = "Court Order Mitchell et al"
			If ga_elig_summ_reason_code = "54" Then ga_elig_summ_reason_info = "Not a GRH Facility"
			If ga_elig_summ_reason_code = "57" Then ga_elig_summ_reason_info = "Undocumented/Inelig Imig"
			If ga_elig_summ_reason_code = "59" Then ga_elig_summ_reason_info = "Imig-status not ver"
			If ga_elig_summ_reason_code = "61" Then ga_elig_summ_reason_info = "No GA Basis or Spouse w/none"
			If ga_elig_summ_reason_code = "62" Then ga_elig_summ_reason_info = "Lump Sum Payment"
			If ga_elig_summ_reason_code = "63" Then ga_elig_summ_reason_info = "Disqualified/Lump Sum"
			If ga_elig_summ_reason_code = "64" Then ga_elig_summ_reason_info = "Failed provide or apply SSN"
			If ga_elig_summ_reason_code = "66" Then ga_elig_summ_reason_info = "Eligible State wide MFIP"
			If ga_elig_summ_reason_code = "96" Then ga_elig_summ_reason_info = "April 2010 Legislation"
			If ga_elig_summ_reason_code = "97" Then ga_elig_summ_reason_info = "GRH Mass Change"
			If ga_elig_summ_reason_code = "98" Then ga_elig_summ_reason_info = "PNA Mass Change"

			ga_elig_summ_worker_message = trim(ga_elig_summ_worker_message)
		End If

		Call back_to_SELF
	end sub
end class

class deny_eligibility_detail

	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found

	public deny_cash_membs_ref_numbs()
	public deny_cash_membs_full_name()
	public deny_cash_membs_request_yn()
	public deny_cash_membs_dwp_test_absence()
	public deny_cash_membs_dwp_test_child_age()
	public deny_cash_membs_dwp_test_citizenship()
	public deny_cash_membs_dwp_test_citizenship_verif()
	public deny_cash_membs_dwp_test_dupl_assist()
	public deny_cash_membs_dwp_test_foster_care()
	public deny_cash_membs_dwp_test_fraud()
	public deny_cash_membs_dwp_test_minor_living_arrangement()
	public deny_cash_membs_dwp_test_post_60_removal()
	public deny_cash_membs_dwp_test_SSI()
	public deny_cash_membs_dwp_test_SSN_coop()
	public deny_cash_membs_dwp_test_Unit_member()
	public deny_cash_membs_dwp_test_unlawful_conduct()
	public deny_cash_membs_mfip_test_absence()
	public deny_cash_membs_mfip_test_child_age()
	public deny_cash_membs_mfip_test_citizenship()
	public deny_cash_membs_mfip_test_citizenship_verif()
	public deny_cash_membs_mfip_test_dupl_assist()
	public deny_cash_membs_mfip_test_foster_care()
	public deny_cash_membs_mfip_test_fraud()
	public deny_cash_membs_mfip_test_minor_living_arrangement()
	public deny_cash_membs_mfip_test_post_60_removal()
	public deny_cash_membs_mfip_test_SSI()
	public deny_cash_membs_mfip_test_SSN_coop()
	public deny_cash_membs_mfip_test_Unit_member()
	public deny_cash_membs_mfip_test_unlawful_conduct()
	public deny_cash_membs_msa_test_absence()
	public deny_cash_membs_msa_test_age()
	public deny_cash_membs_msa_test_basis_of_elig()
	public deny_cash_membs_msa_test_citizenship()
	public deny_cash_membs_msa_test_dupl_assist()
	public deny_cash_membs_msa_test_fail_coop()
	public deny_cash_membs_msa_test_fraud()
	public deny_cash_membs_msa_test_IVE_elig()
	public deny_cash_membs_msa_test_living_arrangment()
	public deny_cash_membs_msa_test_SSI_basis()
	public deny_cash_membs_msa_test_SSN_coop()
	public deny_cash_membs_msa_test_unit_member()
	public deny_cash_membs_msa_test_verif()
	public deny_cash_membs_ga_test_absence()
	public deny_cash_membs_ga_test_basis_of_elig()
	public deny_cash_membs_ga_test_dupl_assist()
	public deny_cash_membs_ga_test_ga_coop()
	public deny_cash_membs_ga_test_imig_status()
	public deny_cash_membs_ga_test_imig_verif()
	public deny_cash_membs_ga_test_IVE_elig()
	public deny_cash_membs_ga_test_lump_sum_payment()
	public deny_cash_membs_ga_test_SSI()
	public deny_cash_membs_ga_test_SSN_coop()
	public deny_cash_membs_ga_test_unit_member()

	public deny_cash_dwp_reason_code
	public deny_cash_mfip_reason_code
	public deny_cash_msa_reason_code
	public deny_cash_ga_reason_code
	public deny_cash_dwp_reason_info
	public deny_cash_mfip_reason_info
	public deny_cash_msa_reason_info
	public deny_cash_ga_reason_info
	public deny_cash_worker_message_one
	public deny_cash_worker_message_two
	public deny_cash_worker_message_three

	public sub read_elig()
		approved_today = False
		approved_version_found = False

		call navigate_to_MAXIS_screen("ELIG", "DENY")
		EMWriteScreen elig_footer_month, 19, 54
		EMWriteScreen elig_footer_year, 19, 57
		Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

			ReDim deny_cash_membs_ref_numbs(0)
			ReDim deny_cash_membs_full_name(0)
			ReDim deny_cash_membs_request_yn(0)
			ReDim deny_cash_membs_dwp_test_absence(0)
			ReDim deny_cash_membs_dwp_test_child_age(0)
			ReDim deny_cash_membs_dwp_test_citizenship(0)
			ReDim deny_cash_membs_dwp_test_citizenship_verif(0)
			ReDim deny_cash_membs_dwp_test_dupl_assist(0)
			ReDim deny_cash_membs_dwp_test_foster_care(0)
			ReDim deny_cash_membs_dwp_test_fraud(0)
			ReDim deny_cash_membs_dwp_test_minor_living_arrangement(0)
			ReDim deny_cash_membs_dwp_test_post_60_removal(0)
			ReDim deny_cash_membs_dwp_test_SSI(0)
			ReDim deny_cash_membs_dwp_test_SSN_coop(0)
			ReDim deny_cash_membs_dwp_test_Unit_member(0)
			ReDim deny_cash_membs_dwp_test_unlawful_conduct(0)
			ReDim deny_cash_membs_mfip_test_absence(0)
			ReDim deny_cash_membs_mfip_test_child_age(0)
			ReDim deny_cash_membs_mfip_test_citizenship(0)
			ReDim deny_cash_membs_mfip_test_citizenship_verif(0)
			ReDim deny_cash_membs_mfip_test_dupl_assist(0)
			ReDim deny_cash_membs_mfip_test_foster_care(0)
			ReDim deny_cash_membs_mfip_test_fraud(0)
			ReDim deny_cash_membs_mfip_test_minor_living_arrangement(0)
			ReDim deny_cash_membs_mfip_test_post_60_removal(0)
			ReDim deny_cash_membs_mfip_test_SSI(0)
			ReDim deny_cash_membs_mfip_test_SSN_coop(0)
			ReDim deny_cash_membs_mfip_test_Unit_member(0)
			ReDim deny_cash_membs_mfip_test_unlawful_conduct(0)
			ReDim deny_cash_membs_msa_test_absence(0)
			ReDim deny_cash_membs_msa_test_age(0)
			ReDim deny_cash_membs_msa_test_basis_of_elig(0)
			ReDim deny_cash_membs_msa_test_citizenship(0)
			ReDim deny_cash_membs_msa_test_dupl_assist(0)
			ReDim deny_cash_membs_msa_test_fail_coop(0)
			ReDim deny_cash_membs_msa_test_fraud(0)
			ReDim deny_cash_membs_msa_test_IVE_elig(0)
			ReDim deny_cash_membs_msa_test_living_arrangment(0)
			ReDim deny_cash_membs_msa_test_SSI_basis(0)
			ReDim deny_cash_membs_msa_test_SSN_coop(0)
			ReDim deny_cash_membs_msa_test_unit_member(0)
			ReDim deny_cash_membs_msa_test_verif(0)
			ReDim deny_cash_membs_ga_test_absence(0)
			ReDim deny_cash_membs_ga_test_basis_of_elig(0)
			ReDim deny_cash_membs_ga_test_dupl_assist(0)
			ReDim deny_cash_membs_ga_test_ga_coop(0)
			ReDim deny_cash_membs_ga_test_imig_status(0)
			ReDim deny_cash_membs_ga_test_imig_verif(0)
			ReDim deny_cash_membs_ga_test_IVE_elig(0)
			ReDim deny_cash_membs_ga_test_lump_sum_payment(0)
			ReDim deny_cash_membs_ga_test_SSI(0)
			ReDim deny_cash_membs_ga_test_SSN_coop(0)
			ReDim deny_cash_membs_ga_test_unit_member(0)

			row = 8
			memb_count = 0
			Do
				ReDim preserve deny_cash_membs_ref_numbs(memb_count)
				ReDim preserve deny_cash_membs_full_name(memb_count)
				ReDim preserve deny_cash_membs_request_yn(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_absence(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_child_age(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_citizenship(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_citizenship_verif(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_dupl_assist(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_foster_care(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_fraud(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_minor_living_arrangement(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_post_60_removal(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_SSI(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_SSN_coop(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_Unit_member(memb_count)
				ReDim preserve deny_cash_membs_dwp_test_unlawful_conduct(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_absence(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_child_age(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_citizenship(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_citizenship_verif(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_dupl_assist(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_foster_care(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_fraud(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_minor_living_arrangement(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_post_60_removal(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_SSI(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_SSN_coop(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_Unit_member(memb_count)
				ReDim preserve deny_cash_membs_mfip_test_unlawful_conduct(memb_count)
				ReDim preserve deny_cash_membs_msa_test_absence(memb_count)
				ReDim preserve deny_cash_membs_msa_test_age(memb_count)
				ReDim preserve deny_cash_membs_msa_test_basis_of_elig(memb_count)
				ReDim preserve deny_cash_membs_msa_test_citizenship(memb_count)
				ReDim preserve deny_cash_membs_msa_test_dupl_assist(memb_count)
				ReDim preserve deny_cash_membs_msa_test_fail_coop(memb_count)
				ReDim preserve deny_cash_membs_msa_test_fraud(memb_count)
				ReDim preserve deny_cash_membs_msa_test_IVE_elig(memb_count)
				ReDim preserve deny_cash_membs_msa_test_living_arrangment(memb_count)
				ReDim preserve deny_cash_membs_msa_test_SSI_basis(memb_count)
				ReDim preserve deny_cash_membs_msa_test_SSN_coop(memb_count)
				ReDim preserve deny_cash_membs_msa_test_unit_member(memb_count)
				ReDim preserve deny_cash_membs_msa_test_verif(memb_count)
				ReDim preserve deny_cash_membs_ga_test_absence(memb_count)
				ReDim preserve deny_cash_membs_ga_test_basis_of_elig(memb_count)
				ReDim preserve deny_cash_membs_ga_test_dupl_assist(memb_count)
				ReDim preserve deny_cash_membs_ga_test_ga_coop(memb_count)
				ReDim preserve deny_cash_membs_ga_test_imig_status(memb_count)
				ReDim preserve deny_cash_membs_ga_test_imig_verif(memb_count)
				ReDim preserve deny_cash_membs_ga_test_IVE_elig(memb_count)
				ReDim preserve deny_cash_membs_ga_test_lump_sum_payment(memb_count)
				ReDim preserve deny_cash_membs_ga_test_SSI(memb_count)
				ReDim preserve deny_cash_membs_ga_test_SSN_coop(memb_count)
				ReDim preserve deny_cash_membs_ga_test_unit_member(memb_count)

				EMReadScreen deny_cash_membs_ref_numbs(memb_count), 2, row, 5
				EMReadScreen deny_cash_membs_full_name(memb_count), 25, row, 11
				EMReadScreen deny_cash_membs_request_yn(memb_count), 1, row, 42

				Call write_value_and_transmit("X", row, 47)		'DWP Tests'
				EMReadScreen deny_cash_membs_dwp_test_absence(memb_count), 				6, 7, 10
				EMReadScreen deny_cash_membs_dwp_test_child_age(memb_count), 			6, 8, 10
				EMReadScreen deny_cash_membs_dwp_test_citizenship(memb_count), 			6, 9, 10
				EMReadScreen deny_cash_membs_dwp_test_citizenship_verif(memb_count), 	6, 10, 10
				EMReadScreen deny_cash_membs_dwp_test_dupl_assist(memb_count), 			6, 11, 10
				EMReadScreen deny_cash_membs_dwp_test_foster_care(memb_count), 			6, 12, 10
				EMReadScreen deny_cash_membs_dwp_test_fraud(memb_count), 				6, 13, 10

				EMReadScreen deny_cash_membs_dwp_test_minor_living_arrangement(memb_count), 6, 7, 42
				EMReadScreen deny_cash_membs_dwp_test_post_60_removal(memb_count), 			6, 8, 42
				EMReadScreen deny_cash_membs_dwp_test_SSI(memb_count), 						6, 9, 42
				EMReadScreen deny_cash_membs_dwp_test_SSN_coop(memb_count), 				6, 10, 42
				EMReadScreen deny_cash_membs_dwp_test_Unit_member(memb_count), 				6, 11, 42
				EMReadScreen deny_cash_membs_dwp_test_unlawful_conduct(memb_count), 		6, 12, 42
				transmit

				Call write_value_and_transmit("X", row, 52)		'MFIP Tests'
				EMReadScreen deny_cash_membs_mfip_test_absence(memb_count), 			6, 7, 10
				EMReadScreen deny_cash_membs_mfip_test_child_age(memb_count), 			6, 8, 10
				EMReadScreen deny_cash_membs_mfip_test_citizenship(memb_count), 		6, 9, 10
				EMReadScreen deny_cash_membs_mfip_test_citizenship_verif(memb_count), 	6, 10, 10
				EMReadScreen deny_cash_membs_mfip_test_dupl_assist(memb_count), 		6, 11, 10
				EMReadScreen deny_cash_membs_mfip_test_foster_care(memb_count), 		6, 12, 10
				EMReadScreen deny_cash_membs_mfip_test_fraud(memb_count), 				6, 13, 10

				EMReadScreen deny_cash_membs_mfip_test_minor_living_arrangement(memb_count), 6, 7, 42
				EMReadScreen deny_cash_membs_mfip_test_post_60_removal(memb_count), 		6, 8, 42
				EMReadScreen deny_cash_membs_mfip_test_SSI(memb_count), 					6, 9, 42
				EMReadScreen deny_cash_membs_mfip_test_SSN_coop(memb_count), 				6, 10, 42
				EMReadScreen deny_cash_membs_mfip_test_Unit_member(memb_count), 			6, 11, 42
				EMReadScreen deny_cash_membs_mfip_test_unlawful_conduct(memb_count), 		6, 12, 42
				transmit

				Call write_value_and_transmit("X", row, 67)		'MSA Tests'
				EMReadScreen deny_cash_membs_msa_test_absence(memb_count), 			6, 7, 10
				EMReadScreen deny_cash_membs_msa_test_age(memb_count), 				6, 8, 10
				EMReadScreen deny_cash_membs_msa_test_basis_of_elig(memb_count), 	6, 9, 10
				EMReadScreen deny_cash_membs_msa_test_citizenship(memb_count), 		6, 10, 10
				EMReadScreen deny_cash_membs_msa_test_dupl_assist(memb_count), 		6, 11, 10
				EMReadScreen deny_cash_membs_msa_test_fail_coop(memb_count), 		6, 12, 10
				EMReadScreen deny_cash_membs_msa_test_fraud(memb_count), 			6, 13, 10

				EMReadScreen deny_cash_membs_msa_test_IVE_elig(memb_count), 		6, 7, 42
				EMReadScreen deny_cash_membs_msa_test_living_arrangment(memb_count), 6, 8, 42
				EMReadScreen deny_cash_membs_msa_test_SSI_basis(memb_count), 		6, 9, 42
				EMReadScreen deny_cash_membs_msa_test_SSN_coop(memb_count), 		6, 10, 42
				EMReadScreen deny_cash_membs_msa_test_unit_member(memb_count), 		6, 11, 42
				EMReadScreen deny_cash_membs_msa_test_verif(memb_count), 			6, 12, 42
				transmit

				Call write_value_and_transmit("X", row, 72)		'GA Tests'
				EMReadScreen deny_cash_membs_ga_test_absence(memb_count), 		6, 7, 10
				EMReadScreen deny_cash_membs_ga_test_basis_of_elig(memb_count), 6, 8, 10
				EMReadScreen deny_cash_membs_ga_test_dupl_assist(memb_count), 	6, 9, 10
				EMReadScreen deny_cash_membs_ga_test_ga_coop(memb_count), 		6, 10, 10
				EMReadScreen deny_cash_membs_ga_test_imig_status(memb_count), 	6, 11, 10
				EMReadScreen deny_cash_membs_ga_test_imig_verif(memb_count), 	6, 12, 10

				EMReadScreen deny_cash_membs_ga_test_IVE_elig(memb_count), 			6, 7, 42
				EMReadScreen deny_cash_membs_ga_test_lump_sum_payment(memb_count), 	6, 8, 42
				EMReadScreen deny_cash_membs_ga_test_SSI(memb_count), 				6, 9, 42
				EMReadScreen deny_cash_membs_ga_test_SSN_coop(memb_count), 			6, 10, 42
				EMReadScreen deny_cash_membs_ga_test_unit_member(memb_count), 		6, 11, 42
				transmit
				row = row + 1
				memb_count = memb_count + 1
				EMReadScreen next_ref_number, 2, row, 5
			Loop until next_ref_number = "  "

			transmit 		'Move to the next panel - CASM

			EMReadScreen deny_cash_dwp_reason_code, 2, 8, 46
			EMReadScreen deny_cash_mfip_reason_code, 2, 9, 46
			EMReadScreen deny_cash_msa_reason_code, 2, 12, 46
			EMReadScreen deny_cash_ga_reason_code, 2, 13, 46

			If deny_cash_dwp_reason_code = "" Then deny_cash_dwp_reason_info = ""
			If deny_cash_dwp_reason_code = "01" Then deny_cash_dwp_reason_info = "No Eligible Child"
			If deny_cash_dwp_reason_code = "02" Then deny_cash_dwp_reason_info = "Application Withdrawn"
			If deny_cash_dwp_reason_code = "03" Then deny_cash_dwp_reason_info = "Initial Income"
			If deny_cash_dwp_reason_code = "04" Then deny_cash_dwp_reason_info = "Assets"
			If deny_cash_dwp_reason_code = "05" Then deny_cash_dwp_reason_info = "Fail To Cooperate"
			If deny_cash_dwp_reason_code = "06" Then deny_cash_dwp_reason_info = "Child Support Disqualification"
			If deny_cash_dwp_reason_code = "07" Then deny_cash_dwp_reason_info = "Employment Services Disqualification"
			If deny_cash_dwp_reason_code = "08" Then deny_cash_dwp_reason_info = "Death"
			If deny_cash_dwp_reason_code = "09" Then deny_cash_dwp_reason_info = "Residence"
			If deny_cash_dwp_reason_code = "10" Then deny_cash_dwp_reason_info = "Transfer of Resources"
			If deny_cash_dwp_reason_code = "11" Then deny_cash_dwp_reason_info = "Verification"
			If deny_cash_dwp_reason_code = "12" Then deny_cash_dwp_reason_info = "Strike"
			If deny_cash_dwp_reason_code = "13" Then deny_cash_dwp_reason_info = "Program Active"
			If deny_cash_dwp_reason_code = "14" Then deny_cash_dwp_reason_info = "4 Month Limit"
			If deny_cash_dwp_reason_code = "15" Then deny_cash_dwp_reason_info = "MFIP Conversion"
			If deny_cash_dwp_reason_code = "23" Then deny_cash_dwp_reason_info = "Duplicate Assistance"
			If deny_cash_dwp_reason_code = "99" Then deny_cash_dwp_reason_info = "PND2 Denial"
			If deny_cash_dwp_reason_code = "TL" Then deny_cash_dwp_reason_info = "TANF Time Limit"

			If deny_cash_mfip_reason_code = "" Then deny_cash_mfip_reason_info = ""
			If deny_cash_mfip_reason_code = "01" Then deny_cash_mfip_reason_info = "No Eligible Child"
			If deny_cash_mfip_reason_code = "02" Then deny_cash_mfip_reason_info = "Application Withdrawn"
			If deny_cash_mfip_reason_code = "03" Then deny_cash_mfip_reason_info = "Initial Income"
			If deny_cash_mfip_reason_code = "04" Then deny_cash_mfip_reason_info = "Monthly Income"
			If deny_cash_mfip_reason_code = "05" Then deny_cash_mfip_reason_info = "Assets"
			If deny_cash_mfip_reason_code = "06" Then deny_cash_mfip_reason_info = "Fail To Cooperate"
			If deny_cash_mfip_reason_code = "07" Then deny_cash_mfip_reason_info = "Fail To Cooperate with IEVS"
			If deny_cash_mfip_reason_code = "08" Then deny_cash_mfip_reason_info = "Death"
			If deny_cash_mfip_reason_code = "09" Then deny_cash_mfip_reason_info = "Residence"
			If deny_cash_mfip_reason_code = "10" Then deny_cash_mfip_reason_info = "Transfer of Resources"
			If deny_cash_mfip_reason_code = "11" Then deny_cash_mfip_reason_info = "Verification"
			If deny_cash_mfip_reason_code = "12" Then deny_cash_mfip_reason_info = "Strike"
			If deny_cash_mfip_reason_code = "13" Then deny_cash_mfip_reason_info = "Fail To File"
			If deny_cash_mfip_reason_code = "14" Then deny_cash_mfip_reason_info = "Program Active"
			If deny_cash_mfip_reason_code = "23" Then deny_cash_mfip_reason_info = "Duplicate Assistance"
			If deny_cash_mfip_reason_code = "24" Then deny_cash_mfip_reason_info = "Minor Living Arrangement"
			If deny_cash_mfip_reason_code = "TL" Then deny_cash_mfip_reason_info = "TANF Time Limit"
			If deny_cash_mfip_reason_code = "33" Then deny_cash_mfip_reason_info = "Diversionary Work Program"
			If deny_cash_mfip_reason_code = "34" Then deny_cash_mfip_reason_info = "Sanction Period"
			If deny_cash_mfip_reason_code = "35" Then deny_cash_mfip_reason_info = "Sanction Date Compliance"
			If deny_cash_mfip_reason_code = "99" Then deny_cash_mfip_reason_info = "PND2 Denial System Entered"

			If deny_cash_msa_reason_code = "" Then deny_cash_msa_reason_info = ""
			If deny_cash_msa_reason_code = "01" Then deny_cash_msa_reason_info = "No Eligible Member"
			If deny_cash_msa_reason_code = "03" Then deny_cash_msa_reason_info = "Verification"
			If deny_cash_msa_reason_code = "08" Then deny_cash_msa_reason_info = "Application Withdrawn"
			If deny_cash_msa_reason_code = "10" Then deny_cash_msa_reason_info = "Residence"
			If deny_cash_msa_reason_code = "11" Then deny_cash_msa_reason_info = "Assets"
			If deny_cash_msa_reason_code = "24" Then deny_cash_msa_reason_info = "Program Active"
			If deny_cash_msa_reason_code = "28" Then deny_cash_msa_reason_info = "Fail To File"
			If deny_cash_msa_reason_code = "29" Then deny_cash_msa_reason_info = "Applicant Eligible"
			If deny_cash_msa_reason_code = "30" Then deny_cash_msa_reason_info = "Prospective Gross Income"
			If deny_cash_msa_reason_code = "31" Then deny_cash_msa_reason_info = "Prospective Net Income"
			If deny_cash_msa_reason_code = "99" Then deny_cash_msa_reason_info = "PND2 Denial System Entered"

			If deny_cash_ga_reason_code = "" Then deny_cash_ga_reason_info = ""
			If deny_cash_ga_reason_code = "01" Then deny_cash_ga_reason_info = "No Eligible Person"
			If deny_cash_ga_reason_code = "02" Then deny_cash_ga_reason_info = "Net Income"
			If deny_cash_ga_reason_code = "03" Then deny_cash_ga_reason_info = "Verification"
			If deny_cash_ga_reason_code = "04" Then deny_cash_ga_reason_info = "Non Cooperation"
			If deny_cash_ga_reason_code = "06" Then deny_cash_ga_reason_info = "Other Benefits"
			If deny_cash_ga_reason_code = "07" Then deny_cash_ga_reason_info = "Address Unknown"
			If deny_cash_ga_reason_code = "08" Then deny_cash_ga_reason_info = "Application Withdrawn"
			If deny_cash_ga_reason_code = "09" Then deny_cash_ga_reason_info = "Client Request"
			If deny_cash_ga_reason_code = "10" Then deny_cash_ga_reason_info = "Residence"
			If deny_cash_ga_reason_code = "11" Then deny_cash_ga_reason_info = "Assets"
			If deny_cash_ga_reason_code = "12" Then deny_cash_ga_reason_info = "Transfer of Resource"
			If deny_cash_ga_reason_code = "14" Then deny_cash_ga_reason_info = "Interim Assistance Agreement"
			If deny_cash_ga_reason_code = "15" Then deny_cash_ga_reason_info = "Out Of County"
			If deny_cash_ga_reason_code = "16" Then deny_cash_ga_reason_info = "Disqualify"
			If deny_cash_ga_reason_code = "17" Then deny_cash_ga_reason_info = "Interview"
			If deny_cash_ga_reason_code = "19" Then deny_cash_ga_reason_info = "Fail to File"
			If deny_cash_ga_reason_code = "21" Then deny_cash_ga_reason_info = "Duplicate Assistance"
			If deny_cash_ga_reason_code = "22" Then deny_cash_ga_reason_info = "Death"
			If deny_cash_ga_reason_code = "23" Then deny_cash_ga_reason_info = "Eligible Other Benefits"
			If deny_cash_ga_reason_code = "26" Then deny_cash_ga_reason_info = "Program Active"
			If deny_cash_ga_reason_code = "29" Then deny_cash_ga_reason_info = "Lump Sum"
			If deny_cash_ga_reason_code = "99" Then deny_cash_ga_reason_info = "PND2 Denial System Entered"

			EMReadScreen deny_cash_worker_message_one, 75, 16, 2
			EMReadScreen deny_cash_worker_message_two, 75, 17, 2
			EMReadScreen deny_cash_worker_message_three, 75, 18, 2

			deny_cash_worker_message_one = trim(deny_cash_worker_message_one)
			deny_cash_worker_message_two = trim(deny_cash_worker_message_two)
			deny_cash_worker_message_three = trim(deny_cash_worker_message_three)
		End If

		Call back_to_SELF
	end sub
end class

class grh_eligibility_detail

	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found
	public er_month
	public hrf_month
	public er_status
	public er_caf_date
	public er_interview_date
	public hrf_status
	public hrf_doc_date

	public grh_elig_memb_ref_numb
	public grh_elig_memb_full_name
	public grh_elig_memb_code
	public grh_elig_memb_info
	public grh_elig_memb_elig_status
	public grh_elig_memb_elig_type_code
	public grh_elig_memb_elig_type_info
	public grh_elig_memb_begin_date

	public grh_elig_case_test_application_withdrawn
	public grh_elig_case_test_pben_coop
	public grh_elig_case_test_elig_thru_other_program
	public grh_elig_case_test_fail_file
	public grh_elig_case_test_placement
	public grh_elig_case_test_state_residence
	public grh_elig_case_test_assets
	public grh_elig_case_test_death_of_applicant
	public grh_elig_case_test_elig_type
	public grh_elig_case_test_income
	public grh_elig_case_test_setting
	public grh_elig_case_test_verif

	public grh_elig_case_test_verif_ACCT
	public grh_elig_case_test_verif_BUSI
	public grh_elig_case_test_verif_CARS
	public grh_elig_case_test_verif_DISA
	public grh_elig_case_test_verif_JOBS
	public grh_elig_case_test_verif_LUMP
	public grh_elig_case_test_verif_MEMB_id
	public grh_elig_case_test_verif_MEMB_dob

	public grh_elig_case_test_verif_PBEN
	public grh_elig_case_test_verif_PACT
	public grh_elig_case_test_verif_RBIC
	public grh_elig_case_test_verif_SECU
	public grh_elig_case_test_verif_STIN
	public grh_elig_case_test_verif_UNEA
	public grh_elig_case_test_verif_TRTX_housing_instability
	public grh_elig_case_test_verif_TRTX_psn_rate_2

	public grh_elig_budg_personal_needs
	public grh_elig_budg_prior_inc_reduce
	public grh_elig_budg_inc_unavail_1st_month
	public grh_elig_budg_total_deductions
	public grh_elig_budg_counted_income
	public grh_elig_budg_total_income

	public grh_elig_budg_SSI_standard_fbr
	public grh_elig_budg_other_countable_PA_grant
	public grh_elig_budg_PASS_disregard
	public grh_elig_budg_MFIP_standard_for_one
	public grh_elig_budg_RSDI_income
	public grh_elig_budg_other_unearned_income
	public grh_elig_budg_earned_income
	public grh_elig_budg_student_EI_disregard
	public grh_elig_budg_standard_EI_disregard
	public grh_elig_budg_earned_income_50_perc_disregard
	public grh_elig_budg_impairment_work_expense
	public grh_elig_budg_child_support_expense
	public grh_elig_budg_child_unmet_need
	public grh_elig_budg_earned_income_subtotal
	public grh_elig_budg_EW_spousal_allocation

	public grh_elig_budg_vendor_number_one
	public grh_elig_budg_total_days_one_one
	public grh_elig_budg_vnd2_rate_limit_one
	public grh_elig_budg_room_board_doc_one
	public grh_elig_budg_total_ssr_rate_one
	public grh_elig_budg_income_test_one
	public grh_elig_payment_grh_state_amount_one
	public grh_elig_payment_county_liability_one
	public grh_elig_payment_total_one
	public grh_elig_payment_amount_already_issued_one

	public grh_elig_budg_vendor_number_two
	public grh_elig_budg_total_days_one_two
	public grh_elig_budg_vnd2_rate_limit_two
	public grh_elig_budg_room_board_doc_two
	public grh_elig_budg_total_ssr_rate_two
	public grh_elig_budg_income_test_two
	public grh_elig_payment_grh_state_amount_two
	public grh_elig_payment_county_liability_two
	public grh_elig_payment_total_two
	public grh_elig_payment_amount_already_issued_two

	public grh_elig_budg_room_board_doc_one_vnd2_days
	public grh_elig_budg_room_board_doc_one_vnd2_rate
	public grh_elig_budg_room_board_doc_one_vnd2_total
	public grh_elig_budg_room_board_doc_one_faci_doc_days
	public grh_elig_budg_room_board_doc_one_faci_doc_rate
	public grh_elig_budg_room_board_doc_one_faci_doc_total
	public grh_elig_budg_room_board_doc_one_total
	public grh_elig_budg_room_board_doc_two_vnd2_days
	public grh_elig_budg_room_board_doc_two_vnd2_rate
	public grh_elig_budg_room_board_doc_two_vnd2_total
	public grh_elig_budg_room_board_doc_two_faci_doc_days
	public grh_elig_budg_room_board_doc_two_faci_doc_rate
	public grh_elig_budg_room_board_doc_two_faci_doc_total
	public grh_elig_budg_room_board_doc_two_total
	public grh_elig_budg_total_ssr_rate_one_days
	public grh_elig_budg_total_ssr_rate_one_rate
	public grh_elig_budg_total_ssr_rate_one_total
	public grh_elig_budg_total_ssr_rate_two_days
	public grh_elig_budg_total_ssr_rate_two_rate
	public grh_elig_budg_total_ssr_rate_two_total
	public grh_elig_payment_county_liability_one_vnd2_co_supp_days
	public grh_elig_payment_county_liability_one_vnd2_co_supp_rate
	public grh_elig_payment_county_liability_one_vnd2_co_supp_total
	public grh_elig_payment_county_liability_one_faci_doc_in_excess_days
	public grh_elig_payment_county_liability_one_faci_doc_in_excess_rate
	public grh_elig_payment_county_liability_one_faci_doc_in_excess_total
	public grh_elig_payment_county_liability_one_total
	public grh_elig_payment_county_liability_two_vnd2_co_supp_days
	public grh_elig_payment_county_liability_two_vnd2_co_supp_rate
	public grh_elig_payment_county_liability_two_vnd2_co_supp_total
	public grh_elig_payment_county_liability_two_faci_doc_in_excess_days
	public grh_elig_payment_county_liability_two_faci_doc_in_excess_rate
	public grh_elig_payment_county_liability_two_faci_doc_in_excess_total
	public grh_elig_payment_county_liability_two_total
	public grh_elig_payment_remaining_income

	public grh_elig_approved_date
	public grh_elig_process_date
	public grh_elig_date_last_approval
	public grh_elig_current_progream_status
	public grh_elig_source_of_info
	public grh_elig_eligibility_result
	' public grh_elig_vendor_number
	public grh_elig_elig_review_date
	public grh_elig_reporting_status
	public grh_elig_responsible_county
	public grh_elig_pre_or_post_pay_one_code
	public grh_elig_pre_or_post_pay_one_info
	public grh_elig_payable_amount_one
	public grh_elig_amount_already_issued_one
	public grh_elig_setting_overpayment_one
	public grh_elig_client_obligation_one
	public grh_elig_pre_or_post_pay_two_code
	public grh_elig_pre_or_post_pay_two_info
	public grh_elig_payable_amount_two
	public grh_elig_amount_already_issued_two
	public grh_elig_setting_overpayment_two
	public grh_elig_client_obligation_two

	public grh_vendor_one_name
	public grh_vendor_one_c_o
	public grh_vendor_one_street_one
	public grh_vendor_one_street_two
	public grh_vendor_one_city
	public grh_vendor_one_state
	public grh_vendor_one_zip
	public grh_vendor_one_grh_yn
	public grh_vendor_one_non_profit_yn
	public grh_vendor_one_phone
	public grh_vendor_one_county
	public grh_vendor_one_status_code
	public grh_vendor_one_status_info
	public grh_vendor_one_incorporated_yn
	public grh_vendor_one_federal_tax_id
	public grh_vendor_one_ssn
	public grh_vendor_one_2nd_address_type_code
	public grh_vendor_one_2nd_address_type_info
	public grh_vendor_one_2nd_address_eff_date
	public grh_vendor_one_2nd_name
	public grh_vendor_one_2nd_c_o
	public grh_vendor_one_2nd_street_one
	public grh_vendor_one_2nd_street_two
	public grh_vendor_one_2nd_city
	public grh_vendor_one_2nd_state
	public grh_vendor_one_2nd_zip
	public grh_vendor_one_direct_deposit_yn
	public grh_vendor_one_merge_vendor_number
	public grh_vendor_one_acct_number_required_yn
	public grh_vendor_one_blocked_county_numbers_list

	public grh_vendor_two_name
	public grh_vendor_two_c_o
	public grh_vendor_two_street_one
	public grh_vendor_two_street_two
	public grh_vendor_two_city
	public grh_vendor_two_state
	public grh_vendor_two_zip
	public grh_vendor_two_grh_yn
	public grh_vendor_two_non_profit_yn
	public grh_vendor_two_phone
	public grh_vendor_two_county
	public grh_vendor_two_status_code
	public grh_vendor_two_status_info
	public grh_vendor_two_incorporated_yn
	public grh_vendor_two_federal_tax_id
	public grh_vendor_two_ssn
	public grh_vendor_two_2nd_address_type_code
	public grh_vendor_two_2nd_address_type_info
	public grh_vendor_two_2nd_address_eff_date
	public grh_vendor_two_2nd_name
	public grh_vendor_two_2nd_c_o
	public grh_vendor_two_2nd_street_one
	public grh_vendor_two_2nd_street_two
	public grh_vendor_two_2nd_city
	public grh_vendor_two_2nd_state
	public grh_vendor_two_2nd_zip
	public grh_vendor_two_direct_deposit_yn
	public grh_vendor_two_merge_vendor_number
	public grh_vendor_two_acct_number_required_yn
	public grh_vendor_two_blocked_county_numbers_list

	public sub read_elig()
		approved_today = False
		approved_version_found = False

		call navigate_to_MAXIS_screen("ELIG", "GRH ")
		EMWriteScreen elig_footer_month, 20, 55
		EMWriteScreen elig_footer_year, 20, 58
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

			EMReadScreen grh_elig_memb_ref_numb, 2, 6, 3
			EMReadScreen grh_elig_memb_full_name, 15, 6, 7
			EMReadScreen grh_elig_memb_code, 1, 6, 24
			If grh_elig_memb_code = "A" Then grh_elig_memb_info  = "Counted"
			EMReadScreen grh_elig_memb_elig_status, 10, 6, 41
			EMReadScreen grh_elig_memb_elig_type_code, 2, 6, 53
			If grh_elig_memb_elig_type_code = "01" Then  grh_elig_memb_elig_type_info = "SSI"
			If grh_elig_memb_elig_type_code = "02" Then  grh_elig_memb_elig_type_info = "MFIP"
			If grh_elig_memb_elig_type_code = "03" Then  grh_elig_memb_elig_type_info = "Blind"
			If grh_elig_memb_elig_type_code = "04" Then  grh_elig_memb_elig_type_info = "Disabled"
			If grh_elig_memb_elig_type_code = "05" Then  grh_elig_memb_elig_type_info = "Aged"
			If grh_elig_memb_elig_type_code = "06" Then  grh_elig_memb_elig_type_info = "Adult"
			If grh_elig_memb_elig_type_code = "07" Then  grh_elig_memb_elig_type_info = "None"
			If grh_elig_memb_elig_type_code = "08" Then  grh_elig_memb_elig_type_info = "Residential Treatment"
			EMReadScreen grh_elig_memb_begin_date, 8, 6, 68

			grh_elig_memb_full_name = trim(grh_elig_memb_full_name)
			grh_elig_memb_elig_status = trim(grh_elig_memb_elig_status)

			grh_elig_memb_begin_date = replace(grh_elig_memb_begin_date, " ", "/")

			EMReadScreen grh_elig_case_test_application_withdrawn, 	6, 8, 8
			EMReadScreen grh_elig_case_test_pben_coop, 				6, 9, 8
			EMReadScreen grh_elig_case_test_elig_thru_other_program, 6, 10, 8
			EMReadScreen grh_elig_case_test_fail_file, 				6, 11, 8
			EMReadScreen grh_elig_case_test_placement, 				6, 12, 8
			EMReadScreen grh_elig_case_test_state_residence, 		6, 13, 8

			EMReadScreen grh_elig_case_test_assets, 				6, 8, 45
			EMReadScreen grh_elig_case_test_death_of_applicant, 	6, 9, 45
			EMReadScreen grh_elig_case_test_elig_type, 				6, 10, 45
			EMReadScreen grh_elig_case_test_income, 				6, 11, 45
			EMReadScreen grh_elig_case_test_setting, 				6, 12, 45
			EMReadScreen grh_elig_case_test_verif, 					6, 13, 45

			grh_elig_case_test_application_withdrawn = trim(grh_elig_case_test_application_withdrawn)
			grh_elig_case_test_pben_coop = trim(grh_elig_case_test_pben_coop)
			grh_elig_case_test_elig_thru_other_program = trim(grh_elig_case_test_elig_thru_other_program)
			grh_elig_case_test_fail_file = trim(grh_elig_case_test_fail_file)
			grh_elig_case_test_placement = trim(grh_elig_case_test_placement)
			grh_elig_case_test_state_residence = trim(grh_elig_case_test_state_residence)

			grh_elig_case_test_assets = trim(grh_elig_case_test_assets)
			grh_elig_case_test_death_of_applicant = trim(grh_elig_case_test_death_of_applicant)
			grh_elig_case_test_elig_type = trim(grh_elig_case_test_elig_type)
			grh_elig_case_test_income = trim(grh_elig_case_test_income)
			grh_elig_case_test_setting = trim(grh_elig_case_test_setting)
			grh_elig_case_test_verif = trim(grh_elig_case_test_verif)

			If grh_elig_case_test_fail_file = "FAILED" Then EMWriteScreen "X", 11, 4
			If grh_elig_case_test_assets = "FAILED" Then EMWriteScreen "X", 8, 41
			If grh_elig_case_test_verif = "FAILED" Then EMWriteScreen "X", 13, 41

			Do
				transmit

				EMReadScreen fail_file_pop_up, 10, 1, 1
				EMReadScreen assets_pop_up, 10, 1, 1
				EMReadScreen verif_pop_up, 10, 1, 1

				If fail_file_pop_up = "" Then
					'TODO - read popup'
				End If

				If assets_pop_up = "" Then
					'TODO - read popup'
				End If

				If verif_pop_up = "" Then
					EMReadScreen grh_elig_case_test_verif_ACCT, 	6, 7, 10
					EMReadScreen grh_elig_case_test_verif_BUSI, 	6, 8, 10
					EMReadScreen grh_elig_case_test_verif_CARS, 	6, 9, 10
					EMReadScreen grh_elig_case_test_verif_DISA, 	6, 10, 10
					EMReadScreen grh_elig_case_test_verif_JOBS, 	6, 11, 10
					EMReadScreen grh_elig_case_test_verif_LUMP, 	6, 12, 10
					EMReadScreen grh_elig_case_test_verif_MEMB_id, 	6, 13, 10
					EMReadScreen grh_elig_case_test_verif_MEMB_dob, 6, 14, 10

					EMReadScreen grh_elig_case_test_verif_PBEN, 	6, 7, 45
					EMReadScreen grh_elig_case_test_verif_PACT, 	6, 8, 45
					EMReadScreen grh_elig_case_test_verif_RBIC, 	6, 9, 45
					EMReadScreen grh_elig_case_test_verif_SECU, 	6, 10, 45
					EMReadScreen grh_elig_case_test_verif_STIN, 	6, 11, 45
					EMReadScreen grh_elig_case_test_verif_UNEA, 	6, 12, 45
					EMReadScreen grh_elig_case_test_verif_TRTX_housing_instability, 6, 13, 45
					EMReadScreen grh_elig_case_test_verif_TRTX_psn_rate_2, 6, 14, 45

					grh_elig_case_test_verif_ACCT = trim(grh_elig_case_test_verif_ACCT)
					grh_elig_case_test_verif_BUSI = trim(grh_elig_case_test_verif_BUSI)
					grh_elig_case_test_verif_CARS = trim(grh_elig_case_test_verif_CARS)
					grh_elig_case_test_verif_DISA = trim(grh_elig_case_test_verif_DISA)
					grh_elig_case_test_verif_JOBS = trim(grh_elig_case_test_verif_JOBS)
					grh_elig_case_test_verif_LUMP = trim(grh_elig_case_test_verif_LUMP)
					grh_elig_case_test_verif_MEMB_id = trim(grh_elig_case_test_verif_MEMB_id)
					grh_elig_case_test_verif_MEMB_dob = trim(grh_elig_case_test_verif_MEMB_dob)

					grh_elig_case_test_verif_PBEN = trim(grh_elig_case_test_verif_PBEN)
					grh_elig_case_test_verif_PACT = trim(grh_elig_case_test_verif_PACT)
					grh_elig_case_test_verif_RBIC = trim(grh_elig_case_test_verif_RBIC)
					grh_elig_case_test_verif_SECU = trim(grh_elig_case_test_verif_SECU)
					grh_elig_case_test_verif_STIN = trim(grh_elig_case_test_verif_STIN)
					grh_elig_case_test_verif_UNEA = trim(grh_elig_case_test_verif_UNEA)
					grh_elig_case_test_verif_TRTX_housing_instability = trim(grh_elig_case_test_verif_TRTX_housing_instability)
					grh_elig_case_test_verif_TRTX_psn_rate_2 = trim(grh_elig_case_test_verif_TRTX_psn_rate_2)
				End If

				EMReadScreen panel_check_GRPB, 4, 3, 51
				EMReadScreen panel_check_GRFB, 4, 3, 47
				If panel_check_GRFB = "GRFB" Then
					skip_budget = True
					Exit Do
				End If
			Loop until panel_check_GRPB = "GRPB"

			If skip_budget = False Then
				If grh_elig_memb_elig_type_info = "SSI" Then
					EMReadScreen grh_elig_budg_SSI_standard_fbr, 		8, 6, 29
					EMReadScreen grh_elig_budg_other_countable_PA_grant, 8, 7, 29

					EMReadScreen grh_elig_budg_PASS_disregard,			8, 10, 29
					EMReadScreen grh_elig_budg_personal_needs, 			8, 11, 29
					EMReadScreen grh_elig_budg_prior_inc_reduce, 		8, 12, 29
					EMReadScreen grh_elig_budg_inc_unavail_1st_month, 	8, 13, 29

					EMReadScreen grh_elig_budg_total_deductions, 		8, 16, 29
					EMReadScreen grh_elig_budg_counted_income,	 		8, 17, 29


					grh_elig_budg_SSI_standard_fbr = replace(grh_elig_budg_SSI_standard_fbr, "_", "")
					grh_elig_budg_other_countable_PA_grant = replace(grh_elig_budg_other_countable_PA_grant, "_", "")

					grh_elig_budg_PASS_disregard = replace(grh_elig_budg_PASS_disregard, "_", "")
					grh_elig_budg_personal_needs = replace(grh_elig_budg_personal_needs, "_", "")
					grh_elig_budg_prior_inc_reduce = replace(grh_elig_budg_prior_inc_reduce, "_", "")
					grh_elig_budg_inc_unavail_1st_month = replace(grh_elig_budg_inc_unavail_1st_month, "_", "")

					grh_elig_budg_total_deductions = trim(grh_elig_budg_total_deductions)
					grh_elig_budg_counted_income = trim(grh_elig_budg_counted_income)
				End If

				If grh_elig_memb_elig_type_info = "MFIP" Then
					EMReadScreen grh_elig_budg_MFIP_standard_for_one, 	8, 6, 31
					EMReadScreen grh_elig_budg_personal_needs, 			8, 9, 31
					EMReadScreen grh_elig_budg_prior_inc_reduce, 		8, 10, 31
					EMReadScreen grh_elig_budg_inc_unavail_1st_month, 	8, 11, 31
					EMReadScreen grh_elig_budg_total_deductions, 		8, 14, 31
					EMReadScreen grh_elig_budg_counted_income,	 		8, 15, 31

					grh_elig_budg_MFIP_standard_for_one = trim(grh_elig_budg_MFIP_standard_for_one)

					grh_elig_budg_personal_needs = replace(grh_elig_budg_personal_needs, "_", "")
					grh_elig_budg_prior_inc_reduce = replace(grh_elig_budg_prior_inc_reduce, "_", "")
					grh_elig_budg_inc_unavail_1st_month = replace(grh_elig_budg_inc_unavail_1st_month, "_", "")

					grh_elig_budg_total_deductions = trim(grh_elig_budg_total_deductions)
					grh_elig_budg_counted_income = trim(grh_elig_budg_counted_income)
				End If

				If grh_elig_memb_elig_type_info = "Blind" or grh_elig_memb_elig_type_info = "Disabled" or grh_elig_memb_elig_type_info = "Aged" or grh_elig_memb_elig_type_info = "Adult" Then
					If grh_elig_memb_elig_type_info = "Aged" or grh_elig_memb_elig_type_info = "Adult" Then
						EMReadScreen grh_elig_budg_RSDI_income, 			8, 8, 27
						EMReadScreen grh_elig_budg_other_unearned_income, 	8, 9, 27
						EMReadScreen grh_elig_budg_earned_income, 			8, 10, 27
						EMReadScreen grh_elig_budg_total_income, 			8, 11, 27
					Else
						EMReadScreen grh_elig_budg_RSDI_income, 			8, 7, 27
						EMReadScreen grh_elig_budg_other_unearned_income, 	8, 8, 27
						EMReadScreen grh_elig_budg_earned_income,	 		8, 9, 27
						EMReadScreen grh_elig_budg_total_income, 			8, 10, 27
					End If

					If grh_elig_memb_elig_type_info = "Adult" Then
						EMReadScreen grh_elig_budg_total_deductions, 	8, 14, 27
						EMReadScreen grh_elig_budg_counted_income, 		8, 15, 27
					Else
						EMReadScreen grh_elig_budg_total_deductions, 	8, 15, 27
						EMReadScreen grh_elig_budg_counted_income, 		8, 16, 27
					End If

					EMReadScreen grh_elig_budg_standard_EI_disregard, 			8, 8, 70

					If grh_elig_memb_elig_type_info = "Blind" Then
						EMReadScreen grh_elig_budg_student_EI_disregard, 			8, 7, 70
						EMReadScreen grh_elig_budg_earned_income_50_perc_disregard, 8, 9, 70
						EMReadScreen grh_elig_budg_impairment_work_expense, 		8, 10, 70
						EMReadScreen grh_elig_budg_personal_needs, 					8, 11, 70
						EMReadScreen grh_elig_budg_child_support_expense, 			8, 12, 70
						EMReadScreen grh_elig_budg_child_unmet_need, 				8, 13, 70
					End If


					If grh_elig_memb_elig_type_info = "Disabled" Then
						EMReadScreen grh_elig_budg_student_EI_disregard, 			8, 7, 70
						EMReadScreen grh_elig_budg_impairment_work_expense, 		8, 9, 70
						EMReadScreen grh_elig_budg_earned_income_50_perc_disregard, 8, 10, 70
						EMReadScreen grh_elig_budg_personal_needs, 					8, 11, 70
						EMReadScreen grh_elig_budg_child_support_expense, 			8, 12, 70
						EMReadScreen grh_elig_budg_child_unmet_need, 				8, 13, 70
					End If

					If grh_elig_memb_elig_type_info = "Aged" Then
						EMReadScreen grh_elig_budg_earned_income_50_perc_disregard, 8, 9, 70
						EMReadScreen grh_elig_budg_personal_needs, 					8, 10, 70
						EMReadScreen grh_elig_budg_child_support_expense, 			8, 11, 70
						EMReadScreen grh_elig_budg_child_unmet_need, 				8, 12, 70
						EMReadScreen grh_elig_budg_EW_spousal_allocation, 			8, 13, 70
					End If

					If grh_elig_memb_elig_type_info = "Adult" Then
						EMReadScreen grh_elig_budg_earned_income_subtotal, 			8, 9, 70
						EMReadScreen grh_elig_budg_earned_income_50_perc_disregard, 8, 10, 70
						EMReadScreen grh_elig_budg_personal_needs, 					8, 11, 70
						EMReadScreen grh_elig_budg_child_support_expense, 			8, 12, 70
						EMReadScreen grh_elig_budg_child_unmet_need, 				8, 13, 70
					End If

					EMReadScreen grh_elig_budg_prior_inc_reduce, 				8, 14, 70
					EMReadScreen grh_elig_budg_inc_unavail_1st_month, 			8, 15, 70

					grh_elig_budg_RSDI_income = replace(grh_elig_budg_RSDI_income, "_", "")
					grh_elig_budg_other_unearned_income = replace(grh_elig_budg_other_unearned_income, "_", "")
					grh_elig_budg_earned_income = replace(grh_elig_budg_earned_income, "_", "")
					grh_elig_budg_total_income = trim(grh_elig_budg_total_income)

					grh_elig_budg_total_deductions = trim(grh_elig_budg_total_deductions)
					grh_elig_budg_counted_income = trim(grh_elig_budg_counted_income)

					grh_elig_budg_student_EI_disregard = replace(grh_elig_budg_student_EI_disregard, "_", "")
					grh_elig_budg_standard_EI_disregard = trim(grh_elig_budg_standard_EI_disregard)
					grh_elig_budg_earned_income_subtotal = trim(grh_elig_budg_earned_income_subtotal)
					grh_elig_budg_earned_income_50_perc_disregard = trim(grh_elig_budg_earned_income_50_perc_disregard)
					grh_elig_budg_impairment_work_expense = replace(grh_elig_budg_impairment_work_expense, "_", "")
					grh_elig_budg_personal_needs = replace(grh_elig_budg_personal_needs, "_", "")
					grh_elig_budg_child_support_expense = replace(grh_elig_budg_child_support_expense, "_", "")
					grh_elig_budg_child_unmet_need = replace(grh_elig_budg_child_unmet_need, "_", "")
					grh_elig_budg_EW_spousal_allocation = replace(grh_elig_budg_EW_spousal_allocation, "_", "")
					grh_elig_budg_prior_inc_reduce = replace(grh_elig_budg_prior_inc_reduce, "_", "")
					grh_elig_budg_inc_unavail_1st_month = replace(grh_elig_budg_inc_unavail_1st_month, "_", "")
				End If

				If grh_elig_memb_elig_type_info = "Residential Treatment" Then
					EMReadScreen grh_elig_budg_total_income, 		8, 12, 25
					EMReadScreen grh_elig_budg_total_deductions, 	8, 15, 25
					EMReadScreen grh_elig_budg_counted_income, 		8, 16, 25

					grh_elig_budg_total_income = trim(grh_elig_budg_total_income)
					grh_elig_budg_total_deductions = trim(grh_elig_budg_total_deductions)
					grh_elig_budg_counted_income = trim(grh_elig_budg_counted_income)
				End If

				transmit 		'go to next panel - GRFB'
			End If

			EMReadScreen grh_elig_budg_vendor_number_one, 	8, 6, 25
			EMReadScreen grh_elig_budg_total_days_one_one, 	8, 7, 25
			EMReadScreen grh_elig_budg_vnd2_rate_limit_one, 8, 8, 25
			EMReadScreen grh_elig_budg_room_board_doc_one, 	8, 9, 25
			' EMReadScreen grh_elig_budg_counted_income, 8, 6, 25
			EMReadScreen grh_elig_budg_total_ssr_rate_one, 	8, 11, 25
			EMReadScreen grh_elig_budg_income_test_one, 	8, 12, 25

			EMReadScreen grh_elig_payment_grh_state_amount_one, 		8, 14, 25
			EMReadScreen grh_elig_payment_county_liability_one, 		8, 15, 25
			' EMReadScreen grh_elig_payment_counted_income, 8, 6, 25
			EMReadScreen grh_elig_payment_total_one, 					8, 17, 25
			EMReadScreen grh_elig_payment_amount_already_issued_one, 	8, 18, 25

			If grh_elig_budg_vendor_number_one = "00000000" Then grh_elig_budg_vendor_number_one = ""
			grh_elig_budg_total_days_one_one = trim(grh_elig_budg_total_days_one_one)
			grh_elig_budg_vnd2_rate_limit_one = trim(grh_elig_budg_vnd2_rate_limit_one)
			grh_elig_budg_room_board_doc_one = trim(grh_elig_budg_room_board_doc_one)
			grh_elig_budg_total_ssr_rate_one = trim(grh_elig_budg_total_ssr_rate_one)
			grh_elig_budg_income_test_one = trim(grh_elig_budg_income_test_one)
			grh_elig_payment_grh_state_amount_one = trim(grh_elig_payment_grh_state_amount_one)
			grh_elig_payment_county_liability_one = trim(grh_elig_payment_county_liability_one)
			grh_elig_payment_total_one = trim(grh_elig_payment_total_one)
			grh_elig_payment_amount_already_issued_one = trim(grh_elig_payment_amount_already_issued_one)

			EMReadScreen grh_elig_budg_vendor_number_two, 	8, 6, 44
			EMReadScreen grh_elig_budg_total_days_one_two, 	8, 7, 44
			EMReadScreen grh_elig_budg_vnd2_rate_limit_two, 8, 8, 44
			EMReadScreen grh_elig_budg_room_board_doc_two, 	8, 9, 44
			' EMReadScreen grh_elig_budg_counted_income, 8, 6, 25
			EMReadScreen grh_elig_budg_total_ssr_rate_two, 	8, 11, 44
			EMReadScreen grh_elig_budg_income_test_two, 	8, 12, 44

			EMReadScreen grh_elig_payment_grh_state_amount_two, 		8, 14, 44
			EMReadScreen grh_elig_payment_county_liability_two, 		8, 15, 44
			' EMReadScreen grh_elig_payment_counted_income, 8, 6, 25
			EMReadScreen grh_elig_payment_total_two, 					8, 17, 44
			EMReadScreen grh_elig_payment_amount_already_issued_two, 	8, 18, 44

			If grh_elig_budg_vendor_number_two = "00000000" Then grh_elig_budg_vendor_number_two = ""
			grh_elig_budg_total_days_one_two = trim(grh_elig_budg_total_days_one_two)
			grh_elig_budg_vnd2_rate_limit_two = trim(grh_elig_budg_vnd2_rate_limit_two)
			grh_elig_budg_room_board_doc_two = trim(grh_elig_budg_room_board_doc_two)
			grh_elig_budg_total_ssr_rate_two = trim(grh_elig_budg_total_ssr_rate_two)
			grh_elig_budg_income_test_two = trim(grh_elig_budg_income_test_two)
			grh_elig_payment_grh_state_amount_two = trim(grh_elig_payment_grh_state_amount_two)
			grh_elig_payment_county_liability_two = trim(grh_elig_payment_county_liability_two)
			grh_elig_payment_total_two = trim(grh_elig_payment_total_two)
			grh_elig_payment_amount_already_issued_two = trim(grh_elig_payment_amount_already_issued_two)

			Call write_value_and_transmit("X", 9, 3)
			EMReadScreen vendor_number_displayed, 8, 16, 26
			vendor_number_displayed = trim(vendor_number_displayed)
			vendor_number_displayed = right("00000000" & vendor_number_displayed, 8)
			If vendor_number_displayed = grh_elig_budg_vendor_number_one Then
				EMReadScreen grh_elig_budg_room_board_doc_one_vnd2_days, 4, 19, 36
				EMReadScreen grh_elig_budg_room_board_doc_one_vnd2_rate, 8, 19, 48
				EMReadScreen grh_elig_budg_room_board_doc_one_vnd2_total, 8, 19, 64

				EMReadScreen grh_elig_budg_room_board_doc_one_faci_doc_days, 4, 20, 36
				EMReadScreen grh_elig_budg_room_board_doc_one_faci_doc_rate, 8, 20, 48
				EMReadScreen grh_elig_budg_room_board_doc_one_faci_doc_total, 8, 20, 64

				EMReadScreen grh_elig_budg_room_board_doc_one_total, 8, 21, 64
			ElseIf vendor_number_displayed = grh_elig_budg_vendor_number_two Then
				EMReadScreen grh_elig_budg_room_board_doc_two_vnd2_days, 4, 19, 36
				EMReadScreen grh_elig_budg_room_board_doc_two_vnd2_rate, 8, 19, 48
				EMReadScreen grh_elig_budg_room_board_doc_two_vnd2_total, 8, 19, 64

				EMReadScreen grh_elig_budg_room_board_doc_two_faci_doc_days, 4, 20, 36
				EMReadScreen grh_elig_budg_room_board_doc_two_faci_doc_rate, 8, 20, 48
				EMReadScreen grh_elig_budg_room_board_doc_two_faci_doc_total, 8, 20, 64

				EMReadScreen grh_elig_budg_room_board_doc_two_total, 8, 21, 64
			End If
			transmit

			EMReadScreen another_vendor_display, 11, 17, 14
			If another_vendor_display = "VENDOR NAME" Then
				EMReadScreen vendor_number_displayed, 8, 16, 26
				vendor_number_displayed = trim(vendor_number_displayed)
				vendor_number_displayed = right("00000000" & vendor_number_displayed, 8)
				If vendor_number_displayed = grh_elig_budg_vendor_number_one Then
					EMReadScreen grh_elig_budg_room_board_doc_one_vnd2_days, 4, 19, 36
					EMReadScreen grh_elig_budg_room_board_doc_one_vnd2_rate, 8, 19, 48
					EMReadScreen grh_elig_budg_room_board_doc_one_vnd2_total, 8, 19, 64

					EMReadScreen grh_elig_budg_room_board_doc_one_faci_doc_days, 4, 20, 36
					EMReadScreen grh_elig_budg_room_board_doc_one_faci_doc_rate, 8, 20, 48
					EMReadScreen grh_elig_budg_room_board_doc_one_faci_doc_total, 8, 20, 64

					EMReadScreen grh_elig_budg_room_board_doc_one_total, 8, 21, 64
				ElseIf vendor_number_displayed = grh_elig_budg_vendor_number_two Then
					EMReadScreen grh_elig_budg_room_board_doc_two_vnd2_days, 4, 19, 36
					EMReadScreen grh_elig_budg_room_board_doc_two_vnd2_rate, 8, 19, 48
					EMReadScreen grh_elig_budg_room_board_doc_two_vnd2_total, 8, 19, 64

					EMReadScreen grh_elig_budg_room_board_doc_two_faci_doc_days, 4, 20, 36
					EMReadScreen grh_elig_budg_room_board_doc_two_faci_doc_rate, 8, 20, 48
					EMReadScreen grh_elig_budg_room_board_doc_two_faci_doc_total, 8, 20, 64

					EMReadScreen grh_elig_budg_room_board_doc_two_total, 8, 21, 64
				End If
				transmit
			End If
			grh_elig_budg_room_board_doc_one_vnd2_days = trim(grh_elig_budg_room_board_doc_one_vnd2_days)
			grh_elig_budg_room_board_doc_one_vnd2_rate = trim(grh_elig_budg_room_board_doc_one_vnd2_rate)
			grh_elig_budg_room_board_doc_one_vnd2_total = trim(grh_elig_budg_room_board_doc_one_vnd2_total)

			grh_elig_budg_room_board_doc_one_faci_doc_days = trim(grh_elig_budg_room_board_doc_one_faci_doc_days)
			grh_elig_budg_room_board_doc_one_faci_doc_rate = trim(grh_elig_budg_room_board_doc_one_faci_doc_rate)
			grh_elig_budg_room_board_doc_one_faci_doc_total = trim(grh_elig_budg_room_board_doc_one_faci_doc_total)

			grh_elig_budg_room_board_doc_one_total = trim(grh_elig_budg_room_board_doc_one_total)


			grh_elig_budg_room_board_doc_two_vnd2_days = trim(grh_elig_budg_room_board_doc_two_vnd2_days)
			grh_elig_budg_room_board_doc_two_vnd2_rate = trim(grh_elig_budg_room_board_doc_two_vnd2_rate)
			grh_elig_budg_room_board_doc_two_vnd2_total = trim(grh_elig_budg_room_board_doc_two_vnd2_total)

			grh_elig_budg_room_board_doc_two_faci_doc_days = trim(grh_elig_budg_room_board_doc_two_faci_doc_days)
			grh_elig_budg_room_board_doc_two_faci_doc_rate = trim(grh_elig_budg_room_board_doc_two_faci_doc_rate)
			grh_elig_budg_room_board_doc_two_faci_doc_total = trim(grh_elig_budg_room_board_doc_two_faci_doc_total)

			grh_elig_budg_room_board_doc_two_total = trim(grh_elig_budg_room_board_doc_two_total)

			Call write_value_and_transmit("X", 11, 3)
			For row = 14 to 19
				EMReadScreen vendor_number_displayed, 8, row, 8
				vendor_number_displayed = trim(vendor_number_displayed)
				vendor_number_displayed = right("00000000" & vendor_number_displayed, 8)
				If vendor_number_displayed = grh_elig_budg_vendor_number_one Then
					EMReadScreen grh_elig_budg_total_ssr_rate_one_days, 5, row, 42
					EMReadScreen grh_elig_budg_total_ssr_rate_one_rate, 9, row, 48
					EMReadScreen grh_elig_budg_total_ssr_rate_one_total, 9, row, 58
				ElseIf vendor_number_displayed = grh_elig_budg_vendor_number_two Then
					EMReadScreen grh_elig_budg_total_ssr_rate_two_days, 5, row, 42
					EMReadScreen grh_elig_budg_total_ssr_rate_two_rate, 9, row, 48
					EMReadScreen grh_elig_budg_total_ssr_rate_two_total, 9, row, 58
				End If
			Next
			grh_elig_budg_total_ssr_rate_one_days = trim(grh_elig_budg_total_ssr_rate_one_days)
			grh_elig_budg_total_ssr_rate_one_rate = trim(grh_elig_budg_total_ssr_rate_one_rate)
			grh_elig_budg_total_ssr_rate_one_total = trim(grh_elig_budg_total_ssr_rate_one_total)
			grh_elig_budg_total_ssr_rate_two_days = trim(grh_elig_budg_total_ssr_rate_two_days)
			grh_elig_budg_total_ssr_rate_two_rate = trim(grh_elig_budg_total_ssr_rate_two_rate)
			grh_elig_budg_total_ssr_rate_two_total = trim(grh_elig_budg_total_ssr_rate_two_total)
			transmit

			Call write_value_and_transmit("X", 15, 3)

			EMReadScreen vendor_number_displayed, 8, 2, 26
			vendor_number_displayed = trim(vendor_number_displayed)
			vendor_number_displayed = right("00000000" & vendor_number_displayed, 8)
			If vendor_number_displayed = grh_elig_budg_vendor_number_one Then
				EMReadScreen grh_elig_payment_county_liability_one_vnd2_co_supp_days, 4, 5, 42
				EMReadScreen grh_elig_payment_county_liability_one_vnd2_co_supp_rate, 8, 5, 55
				EMReadScreen grh_elig_payment_county_liability_one_vnd2_co_supp_total, 8, 5, 68

				EMReadScreen grh_elig_payment_county_liability_one_faci_doc_in_excess_days, 4, 6, 42
				EMReadScreen grh_elig_payment_county_liability_one_faci_doc_in_excess_rate, 8, 6, 55
				EMReadScreen grh_elig_payment_county_liability_one_faci_doc_in_excess_total, 8, 6, 68

				EMReadScreen grh_elig_payment_county_liability_one_total, 8, 7, 68
			ElseIf vendor_number_displayed = grh_elig_budg_vendor_number_two Then
				EMReadScreen grh_elig_payment_county_liability_two_vnd2_co_supp_days, 4, 5, 42
				EMReadScreen grh_elig_payment_county_liability_two_vnd2_co_supp_rate, 8, 5, 55
				EMReadScreen grh_elig_payment_county_liability_two_vnd2_co_supp_total, 8, 5, 68

				EMReadScreen grh_elig_payment_county_liability_two_faci_doc_in_excess_days, 4, 6, 42
				EMReadScreen grh_elig_payment_county_liability_two_faci_doc_in_excess_rate, 8, 6, 55
				EMReadScreen grh_elig_payment_county_liability_two_faci_doc_in_excess_total, 8, 6, 68

				EMReadScreen grh_elig_payment_county_liability_two_total, 8, 7, 68
			End If
			transmit

			EMReadScreen another_vendor_display, 11, 3, 14
			If another_vendor_display = "Vendor Name]" Then
				EMReadScreen vendor_number_displayed, 8, 2, 26
				vendor_number_displayed = trim(vendor_number_displayed)
				vendor_number_displayed = right("00000000" & vendor_number_displayed, 8)
				If vendor_number_displayed = grh_elig_budg_vendor_number_one Then
					EMReadScreen grh_elig_payment_county_liability_one_vnd2_co_supp_days, 4, 5, 42
					EMReadScreen grh_elig_payment_county_liability_one_vnd2_co_supp_rate, 8, 5, 55
					EMReadScreen grh_elig_payment_county_liability_one_vnd2_co_supp_total, 8, 5, 68

					EMReadScreen grh_elig_payment_county_liability_one_faci_doc_in_excess_days, 4, 6, 42
					EMReadScreen grh_elig_payment_county_liability_one_faci_doc_in_excess_rate, 8, 6, 55
					EMReadScreen grh_elig_payment_county_liability_one_faci_doc_in_excess_total, 8, 6, 68

					EMReadScreen grh_elig_payment_county_liability_one_total, 8, 6, 68
				ElseIf vendor_number_displayed = grh_elig_budg_vendor_number_two Then
					EMReadScreen grh_elig_payment_county_liability_two_vnd2_co_supp_days, 4, 5, 42
					EMReadScreen grh_elig_payment_county_liability_two_vnd2_co_supp_rate, 8, 5, 55
					EMReadScreen grh_elig_payment_county_liability_two_vnd2_co_supp_total, 8, 5, 68

					EMReadScreen grh_elig_payment_county_liability_two_faci_doc_in_excess_days, 4, 6, 42
					EMReadScreen grh_elig_payment_county_liability_two_faci_doc_in_excess_rate, 8, 6, 55
					EMReadScreen grh_elig_payment_county_liability_two_faci_doc_in_excess_total, 8, 6, 68

					EMReadScreen grh_elig_payment_county_liability_two_total, 8, 7, 68
				End If
				transmit
			End If
			grh_elig_payment_county_liability_one_vnd2_co_supp_days = trim(grh_elig_payment_county_liability_one_vnd2_co_supp_days)
			grh_elig_payment_county_liability_one_vnd2_co_supp_rate = trim(grh_elig_payment_county_liability_one_vnd2_co_supp_rate)
			grh_elig_payment_county_liability_one_vnd2_co_supp_total = trim(grh_elig_payment_county_liability_one_vnd2_co_supp_total)

			grh_elig_payment_county_liability_one_faci_doc_in_excess_days = trim(grh_elig_payment_county_liability_one_faci_doc_in_excess_days)
			grh_elig_payment_county_liability_one_faci_doc_in_excess_rate = trim(grh_elig_payment_county_liability_one_faci_doc_in_excess_rate)
			grh_elig_payment_county_liability_one_faci_doc_in_excess_total = trim(grh_elig_payment_county_liability_one_faci_doc_in_excess_total)

			grh_elig_payment_county_liability_one_total = trim(grh_elig_payment_county_liability_one_total)


			grh_elig_payment_county_liability_two_vnd2_co_supp_days = trim(grh_elig_payment_county_liability_two_vnd2_co_supp_days)
			grh_elig_payment_county_liability_two_vnd2_co_supp_rate = trim(grh_elig_payment_county_liability_two_vnd2_co_supp_rate)
			grh_elig_payment_county_liability_two_vnd2_co_supp_total = trim(grh_elig_payment_county_liability_two_vnd2_co_supp_total)

			grh_elig_payment_county_liability_two_faci_doc_in_excess_days = trim(grh_elig_payment_county_liability_two_faci_doc_in_excess_days)
			grh_elig_payment_county_liability_two_faci_doc_in_excess_rate = trim(grh_elig_payment_county_liability_two_faci_doc_in_excess_rate)
			grh_elig_payment_county_liability_two_faci_doc_in_excess_total = trim(grh_elig_payment_county_liability_two_faci_doc_in_excess_total)

			grh_elig_payment_county_liability_two_total = trim(grh_elig_payment_county_liability_two_total)

			transmit

			Call write_value_and_transmit("X", 16, 3)
			EMReadScreen grh_elig_payment_remaining_income, 9, 4, 53
			grh_elig_payment_remaining_income = trim(grh_elig_payment_remaining_income)
			transmit

			transmit 		'go to next panel - GRSM

			EMReadScreen grh_elig_approved_date, 8, 3, 14
			EMReadScreen grh_elig_process_date, 8, 2, 72

			EMReadScreen grh_elig_date_last_approval, 		8, 5, 31
			EMReadScreen grh_elig_current_progream_status, 	10, 6, 31
			EMReadScreen grh_elig_source_of_info, 			4, 7, 31
			EMReadScreen grh_elig_eligibility_result, 		10, 8, 31

			EMReadScreen grh_elig_elig_review_date, 	8, 6, 69
			EMReadScreen grh_elig_reporting_status, 	8, 7, 69
			EMReadScreen grh_elig_responsible_county, 	2, 8, 69

			grh_elig_current_progream_status = trim(grh_elig_current_progream_status)
			grh_elig_eligibility_result = trim(grh_elig_eligibility_result)

			grh_elig_elig_review_date = replace(grh_elig_elig_review_date, " ", "/")
			grh_elig_reporting_status = trim(grh_elig_reporting_status)

			' EMReadScreen grh_elig_vendor_number, 		8, 10, 32
			EMReadScreen grh_elig_pre_or_post_pay_one_code, 2, 11, 38
			EMReadScreen grh_elig_payable_amount_one, 		9, 12, 31

			EMReadScreen grh_elig_amount_already_issued_one, 9, 13, 31
			EMReadScreen grh_elig_setting_overpayment_one, 	9, 16, 31
			EMReadScreen grh_elig_client_obligation_one, 	9, 17, 31

			If grh_elig_pre_or_post_pay_one_code = "07" Then grh_elig_pre_or_post_pay_one_info = "Pre-pay"
			If grh_elig_pre_or_post_pay_one_code = "08" Then grh_elig_pre_or_post_pay_one_info = "Post-pay Hold"
			If grh_elig_pre_or_post_pay_one_code = "20" Then grh_elig_pre_or_post_pay_one_info = "Release Post-pay"
			grh_elig_payable_amount_one = trim(grh_elig_payable_amount_one)
			grh_elig_amount_already_issued_one = trim(grh_elig_amount_already_issued_one)
			grh_elig_setting_overpayment_one = trim(grh_elig_setting_overpayment_one)
			grh_elig_client_obligation_one = trim(grh_elig_client_obligation_one)

			EMReadScreen grh_elig_pre_or_post_pay_two_code, 2, 11, 57
			EMReadScreen grh_elig_payable_amount_two, 		9, 12, 50

			EMReadScreen grh_elig_amount_already_issued_two, 9, 13, 50
			EMReadScreen grh_elig_setting_overpayment_two, 	9, 16, 50
			EMReadScreen grh_elig_client_obligation_two, 	9, 17, 50

			If grh_elig_pre_or_post_pay_two_code = "07" Then grh_elig_pre_or_post_pay_two_info = "Pre-pay"
			If grh_elig_pre_or_post_pay_two_code = "08" Then grh_elig_pre_or_post_pay_two_info = "Post-pay Hold"
			If grh_elig_pre_or_post_pay_two_code = "20" Then grh_elig_pre_or_post_pay_two_info = "Release Post-pay"
			grh_elig_payable_amount_two = trim(grh_elig_payable_amount_two)
			grh_elig_amount_already_issued_two = trim(grh_elig_amount_already_issued_two)
			grh_elig_setting_overpayment_two = trim(grh_elig_setting_overpayment_two)
			grh_elig_client_obligation_two = trim(grh_elig_client_obligation_two)

			call back_to_SELF

			Call navigate_to_MAXIS_screen("MONY", "VNDS")

			If grh_elig_budg_vendor_number_one <> "" Then
				Call write_value_and_transmit(grh_elig_budg_vendor_number_one, 4, 59)
				EMReadScreen grh_vendor_one_name, 					30, 3, 15
				EMReadScreen grh_vendor_one_c_o, 					30, 4, 15
				EMReadScreen grh_vendor_one_street_one, 			22, 5, 15
				EMReadScreen grh_vendor_one_street_two, 			22, 6, 15
				EMReadScreen grh_vendor_one_city, 					15, 7, 15
				EMReadScreen grh_vendor_one_state, 					2, 7, 36
				EMReadScreen grh_vendor_one_zip, 					10, 7, 46
				EMReadScreen grh_vendor_one_grh_yn, 				1, 4, 57
				EMReadScreen grh_vendor_one_non_profit_yn, 			1, 4, 78
				EMReadScreen grh_vendor_one_phone, 					16, 6, 54
				grh_vendor_one_phone = "(" & replace(replace(grh_vendor_one_phone, " )  ", ")"), "  ", "-")
				EMReadScreen grh_vendor_one_county, 				2, 7, 61
				EMReadScreen grh_vendor_one_status_code, 			1, 16, 15
				If grh_vendor_one_status_code = "A" Then grh_vendor_one_status_info = "Active"
				If grh_vendor_one_status_code = "D" Then grh_vendor_one_status_info = "Delete"
				If grh_vendor_one_status_code = "M" Then grh_vendor_one_status_info = "Merged"
				If grh_vendor_one_status_code = "P" Then grh_vendor_one_status_info = "Pending"
				If grh_vendor_one_status_code = "T" Then grh_vendor_one_status_info = "Terminated"
				EMReadScreen grh_vendor_one_incorporated_yn, 		1, 9, 22
				EMReadScreen grh_vendor_one_federal_tax_id, 		9, 9, 41
				EMReadScreen grh_vendor_one_ssn, 					11, 9, 61
				If grh_vendor_one_ssn = "___ __ ____" Then grh_vendor_one_ssn = ""
				grh_vendor_one_ssn = replace(grh_vendor_one_ssn, " ", "-")
				EMReadScreen grh_vendor_one_2nd_address_type_code, 	1, 10, 22
				If grh_vendor_one_2nd_address_type_code = "1" Then grh_vendor_one_2nd_address_type_info = "Mailing Address"
				If grh_vendor_one_2nd_address_type_code = "2" Then grh_vendor_one_2nd_address_type_info = "Court Order"
				EMReadScreen grh_vendor_one_2nd_address_eff_date, 	8, 11, 15
				If grh_vendor_one_2nd_address_eff_date = "__ __ __" Then grh_vendor_one_2nd_address_eff_date = ""
				grh_vendor_one_2nd_address_eff_date = replace(grh_vendor_one_2nd_address_eff_date, " ", "/")
				EMReadScreen grh_vendor_one_2nd_name, 				30, 11, 15
				EMReadScreen grh_vendor_one_2nd_c_o, 				30, 12, 15
				EMReadScreen grh_vendor_one_2nd_street_one, 		22, 13, 15
				EMReadScreen grh_vendor_one_2nd_street_two, 		22, 14, 15
				EMReadScreen grh_vendor_one_2nd_city, 				15, 15, 15
				EMReadScreen grh_vendor_one_2nd_state, 				2, 15, 35
				EMReadScreen grh_vendor_one_2nd_zip, 				10, 15, 44
				EMReadScreen grh_vendor_one_direct_deposit_yn, 		1, 12, 76
				EMReadScreen grh_vendor_one_merge_vendor_number, 	8, 16, 38
				EMReadScreen grh_vendor_one_acct_number_required_yn, 1, 17, 74
				EMReadScreen grh_vendor_one_blocked_county_numbers_list, 29, 18, 23

				grh_vendor_one_name = replace(grh_vendor_one_name, "_", "")
				grh_vendor_one_c_o = replace(grh_vendor_one_c_o, "_", "")
				grh_vendor_one_street_one = replace(grh_vendor_one_street_one, "_", "")
				grh_vendor_one_street_two = replace(grh_vendor_one_street_two, "_", "")
				grh_vendor_one_city = replace(grh_vendor_one_city, "_", "")
				grh_vendor_one_zip = trim(grh_vendor_one_zip)
				grh_vendor_one_zip = replace(grh_vendor_one_zip, " ", "-")

				grh_vendor_one_federal_tax_id = replace(grh_vendor_one_federal_tax_id, "_", "")

				grh_vendor_one_2nd_name = replace(grh_vendor_one_2nd_name, "_", "")
				grh_vendor_one_2nd_c_o = replace(grh_vendor_one_2nd_c_o, "_", "")
				grh_vendor_one_2nd_street_one = replace(grh_vendor_one_2nd_street_one, "_", "")
				grh_vendor_one_2nd_street_two = replace(grh_vendor_one_2nd_street_two, "_", "")
				grh_vendor_one_2nd_city = replace(grh_vendor_one_2nd_city, "_", "")
				grh_vendor_one_2nd_zip = replace(grh_vendor_one_2nd_zip, "_", "")
				grh_vendor_one_2nd_zip = trim(grh_vendor_one_2nd_zip)
				grh_vendor_one_2nd_zip = replace(grh_vendor_one_2nd_zip, " ", "-")

				grh_vendor_one_merge_vendor_number = replace(grh_vendor_one_merge_vendor_number, "_", "")
				grh_vendor_one_acct_number_required_yn = replace(grh_vendor_one_acct_number_required_yn, "_", "")

				grh_vendor_one_blocked_county_numbers_list = replace(grh_vendor_one_blocked_county_numbers_list, "_", "")
				grh_vendor_one_blocked_county_numbers_list = trim(grh_vendor_one_blocked_county_numbers_list)

				transmit
				EMReadScreen grh_vendor_one_current_rate_period_code, 1, 6, 24

				If grh_vendor_one_current_rate_period_code = "1" Then grh_vendor_one_current_rate_period_info = "Calendar Year"
				If grh_vendor_one_current_rate_period_code = "2" Then grh_vendor_one_current_rate_period_info = "Fiscal Year"
				If grh_vendor_one_current_rate_period_code = "3" Then grh_vendor_one_current_rate_period_info = "Federal Fiscal Year"
				If grh_vendor_one_current_rate_period_code = "4" Then grh_vendor_one_current_rate_period_info = "Other"

				EMReadScreen grh_vendor_one_rate_from_date, 7, 6, 47
				EMReadScreen grh_vendor_one_rate_to_date, 7, 6, 61
				EMReadScreen grh_vendor_one_initial_rate_date, 10, 7, 21
				EMReadScreen grh_vendor_one_NPI_number, 10, 7, 41
				EMReadScreen grh_vendor_one_family_foster_care_yn, 1, 8, 69
				EMReadScreen grh_vendor_one_rate_limit, 8, 9, 24
				EMReadScreen grh_vendor_one_exempt_reason_code, 2, 9, 69
				If grh_vendor_one_exempt_reason_code = "01" Then grh_vendor_one_exempt_reason_info = "Andrew Board & Care"
				If grh_vendor_one_exempt_reason_code = "04" Then grh_vendor_one_exempt_reason_info = "Aldrich"
				If grh_vendor_one_exempt_reason_code = "05" Then grh_vendor_one_exempt_reason_info = "Amy Johnson"
				If grh_vendor_one_exempt_reason_code = "09" Then grh_vendor_one_exempt_reason_info = "Quinlan Home"
				If grh_vendor_one_exempt_reason_code = "10" Then grh_vendor_one_exempt_reason_info = "Revere Home"
				If grh_vendor_one_exempt_reason_code = "11" Then grh_vendor_one_exempt_reason_info = "River Oaks"
				If grh_vendor_one_exempt_reason_code = "12" Then grh_vendor_one_exempt_reason_info = "Special Srvices"
				If grh_vendor_one_exempt_reason_code = "14" Then grh_vendor_one_exempt_reason_info = "Albert Lea"
				If grh_vendor_one_exempt_reason_code = "15" Then grh_vendor_one_exempt_reason_info = "Metro Demo"
				If grh_vendor_one_exempt_reason_code = "16" Then grh_vendor_one_exempt_reason_info = "Broadway"
				If grh_vendor_one_exempt_reason_code = "18" Then grh_vendor_one_exempt_reason_info = "Murphy's Board and Care"
				If grh_vendor_one_exempt_reason_code = "25" Then grh_vendor_one_exempt_reason_info = "Valley Home"
				If grh_vendor_one_exempt_reason_code = "26" Then grh_vendor_one_exempt_reason_info = "LTH Supportive Housing"
				If grh_vendor_one_exempt_reason_code = "27" Then grh_vendor_one_exempt_reason_info = "Boarding Care Home"
				If grh_vendor_one_exempt_reason_code = "28" Then grh_vendor_one_exempt_reason_info = "Banked Bed"
				If grh_vendor_one_exempt_reason_code = "29" Then grh_vendor_one_exempt_reason_info = "Tribe Certified Housing"

				EMReadScreen grh_vendor_one_DHS_license_1_code, 2, 10, 24
				If grh_vendor_one_DHS_license_1_code = "__" Then grh_vendor_one_DHS_license_1_info = ""
				If grh_vendor_one_DHS_license_1_code = "01" Then grh_vendor_one_DHS_license_1_info = "SILS- Developmental Disabled Rule 18"
				If grh_vendor_one_DHS_license_1_code = "02" Then grh_vendor_one_DHS_license_1_info = "Developmentaly Diabled Rule 34"
				If grh_vendor_one_DHS_license_1_code = "03" Then grh_vendor_one_DHS_license_1_info = "Adult Mentally Ill Rule 36"
				If grh_vendor_one_DHS_license_1_code = "04" Then grh_vendor_one_DHS_license_1_info = "Adult Foster Care Rule 203"
				If grh_vendor_one_DHS_license_1_code = "05" Then grh_vendor_one_DHS_license_1_info = "Mentally Retarded Waiver Rule 42"
				If grh_vendor_one_DHS_license_1_code = "06" Then grh_vendor_one_DHS_license_1_info = "Pregnant Woman Shelter Rule 6"
				If grh_vendor_one_DHS_license_1_code = "07" Then grh_vendor_one_DHS_license_1_info = "Other DHS license"
				If grh_vendor_one_DHS_license_1_code = "08" Then grh_vendor_one_DHS_license_1_info = "No DHS License"
				If grh_vendor_one_DHS_license_1_code = "09" Then grh_vendor_one_DHS_license_1_info = "Physical Handicap Rule 80"
				If grh_vendor_one_DHS_license_1_code = "10" Then grh_vendor_one_DHS_license_1_info = "Child Foster Care Rules 1 & 8"
				If grh_vendor_one_DHS_license_1_code = "11" Then grh_vendor_one_DHS_license_1_info = "Chemical Dependancy Rule 35"
				EMReadScreen grh_vendor_one_DHS_license_2_code, 2, 10, 27
				If grh_vendor_one_DHS_license_2_code = "__" Then grh_vendor_one_DHS_license_2_info = ""
				If grh_vendor_one_DHS_license_2_code = "01" Then grh_vendor_one_DHS_license_2_info = "SILS- Developmental Disabled Rule 18"
				If grh_vendor_one_DHS_license_2_code = "02" Then grh_vendor_one_DHS_license_2_info = "Developmentaly Diabled Rule 34"
				If grh_vendor_one_DHS_license_2_code = "03" Then grh_vendor_one_DHS_license_2_info = "Adult Mentally Ill Rule 36"
				If grh_vendor_one_DHS_license_2_code = "04" Then grh_vendor_one_DHS_license_2_info = "Adult Foster Care Rule 203"
				If grh_vendor_one_DHS_license_2_code = "05" Then grh_vendor_one_DHS_license_2_info = "Mentally Retarded Waiver Rule 42"
				If grh_vendor_one_DHS_license_2_code = "06" Then grh_vendor_one_DHS_license_2_info = "Pregnant Woman Shelter Rule 6"
				If grh_vendor_one_DHS_license_2_code = "07" Then grh_vendor_one_DHS_license_2_info = "Other DHS license"
				If grh_vendor_one_DHS_license_2_code = "08" Then grh_vendor_one_DHS_license_2_info = "No DHS License"
				If grh_vendor_one_DHS_license_2_code = "09" Then grh_vendor_one_DHS_license_2_info = "Physical Handicap Rule 80"
				If grh_vendor_one_DHS_license_2_code = "10" Then grh_vendor_one_DHS_license_2_info = "Child Foster Care Rules 1 & 8"
				If grh_vendor_one_DHS_license_2_code = "11" Then grh_vendor_one_DHS_license_2_info = "Chemical Dependancy Rule 35"
				EMReadScreen grh_vendor_one_DHS_license_3_code, 2, 10, 30
				If grh_vendor_one_DHS_license_3_code = "__" Then grh_vendor_one_DHS_license_3_info = ""
				If grh_vendor_one_DHS_license_3_code = "01" Then grh_vendor_one_DHS_license_3_info = "SILS- Developmental Disabled Rule 18"
				If grh_vendor_one_DHS_license_3_code = "02" Then grh_vendor_one_DHS_license_3_info = "Developmentaly Diabled Rule 34"
				If grh_vendor_one_DHS_license_3_code = "03" Then grh_vendor_one_DHS_license_3_info = "Adult Mentally Ill Rule 36"
				If grh_vendor_one_DHS_license_3_code = "04" Then grh_vendor_one_DHS_license_3_info = "Adult Foster Care Rule 203"
				If grh_vendor_one_DHS_license_3_code = "05" Then grh_vendor_one_DHS_license_3_info = "Mentally Retarded Waiver Rule 42"
				If grh_vendor_one_DHS_license_3_code = "06" Then grh_vendor_one_DHS_license_3_info = "Pregnant Woman Shelter Rule 6"
				If grh_vendor_one_DHS_license_3_code = "07" Then grh_vendor_one_DHS_license_3_info = "Other DHS license"
				If grh_vendor_one_DHS_license_3_code = "08" Then grh_vendor_one_DHS_license_3_info = "No DHS License"
				If grh_vendor_one_DHS_license_3_code = "09" Then grh_vendor_one_DHS_license_3_info = "Physical Handicap Rule 80"
				If grh_vendor_one_DHS_license_3_code = "10" Then grh_vendor_one_DHS_license_3_info = "Child Foster Care Rules 1 & 8"
				If grh_vendor_one_DHS_license_3_code = "11" Then grh_vendor_one_DHS_license_3_info = "Chemical Dependancy Rule 35"

				EMReadScreen grh_vendor_one_health_dept_license_1_code, 2, 10, 69
				If grh_vendor_one_health_dept_license_1_code = "__" Then grh_vendor_one_health_dept_license_1_info = ""
				If grh_vendor_one_health_dept_license_1_code = "01" Then grh_vendor_one_health_dept_license_1_info = "Nursing Home"
				If grh_vendor_one_health_dept_license_1_code = "02" Then grh_vendor_one_health_dept_license_1_info = "Boarding Care Home"
				If grh_vendor_one_health_dept_license_1_code = "03" Then grh_vendor_one_health_dept_license_1_info = "Supervised Living Facility"
				If grh_vendor_one_health_dept_license_1_code = "04" Then grh_vendor_one_health_dept_license_1_info = "Board and Lodging"
				If grh_vendor_one_health_dept_license_1_code = "05" Then grh_vendor_one_health_dept_license_1_info = "Hotal/Restaurant"
				If grh_vendor_one_health_dept_license_1_code = "06" Then grh_vendor_one_health_dept_license_1_info = "Board & Lodge with Special Services"
				If grh_vendor_one_health_dept_license_1_code = "07" Then grh_vendor_one_health_dept_license_1_info = "Tribal License"
				If grh_vendor_one_health_dept_license_1_code = "08" Then grh_vendor_one_health_dept_license_1_info = "Metro Demo"
				If grh_vendor_one_health_dept_license_1_code = "09" Then grh_vendor_one_health_dept_license_1_info = "Housing with Services"
				If grh_vendor_one_health_dept_license_1_code = "10" Then grh_vendor_one_health_dept_license_1_info = "Supportive Housing"
				EMReadScreen grh_vendor_one_health_dept_license_2_code, 2, 10, 72
				If grh_vendor_one_health_dept_license_2_code = "__" Then grh_vendor_one_health_dept_license_2_info = ""
				If grh_vendor_one_health_dept_license_2_code = "01" Then grh_vendor_one_health_dept_license_2_info = "Nursing Home"
				If grh_vendor_one_health_dept_license_2_code = "02" Then grh_vendor_one_health_dept_license_2_info = "Boarding Care Home"
				If grh_vendor_one_health_dept_license_2_code = "03" Then grh_vendor_one_health_dept_license_2_info = "Supervised Living Facility"
				If grh_vendor_one_health_dept_license_2_code = "04" Then grh_vendor_one_health_dept_license_2_info = "Board and Lodging"
				If grh_vendor_one_health_dept_license_2_code = "05" Then grh_vendor_one_health_dept_license_2_info = "Hotal/Restaurant"
				If grh_vendor_one_health_dept_license_2_code = "06" Then grh_vendor_one_health_dept_license_2_info = "Board & Lodge with Special Services"
				If grh_vendor_one_health_dept_license_2_code = "07" Then grh_vendor_one_health_dept_license_2_info = "Tribal License"
				If grh_vendor_one_health_dept_license_2_code = "08" Then grh_vendor_one_health_dept_license_2_info = "Metro Demo"
				If grh_vendor_one_health_dept_license_2_code = "09" Then grh_vendor_one_health_dept_license_2_info = "Housing with Services"
				If grh_vendor_one_health_dept_license_2_code = "10" Then grh_vendor_one_health_dept_license_2_info = "Supportive Housing"
				EMReadScreen grh_vendor_one_health_dept_license_3_code, 2, 10, 75
				If grh_vendor_one_health_dept_license_3_code = "__" Then grh_vendor_one_health_dept_license_3_info = ""
				If grh_vendor_one_health_dept_license_3_code = "01" Then grh_vendor_one_health_dept_license_3_info = "Nursing Home"
				If grh_vendor_one_health_dept_license_3_code = "02" Then grh_vendor_one_health_dept_license_3_info = "Boarding Care Home"
				If grh_vendor_one_health_dept_license_3_code = "03" Then grh_vendor_one_health_dept_license_3_info = "Supervised Living Facility"
				If grh_vendor_one_health_dept_license_3_code = "04" Then grh_vendor_one_health_dept_license_3_info = "Board and Lodging"
				If grh_vendor_one_health_dept_license_3_code = "05" Then grh_vendor_one_health_dept_license_3_info = "Hotal/Restaurant"
				If grh_vendor_one_health_dept_license_3_code = "06" Then grh_vendor_one_health_dept_license_3_info = "Board & Lodge with Special Services"
				If grh_vendor_one_health_dept_license_3_code = "07" Then grh_vendor_one_health_dept_license_3_info = "Tribal License"
				If grh_vendor_one_health_dept_license_3_code = "08" Then grh_vendor_one_health_dept_license_3_info = "Metro Demo"
				If grh_vendor_one_health_dept_license_3_code = "09" Then grh_vendor_one_health_dept_license_3_info = "Housing with Services"
				If grh_vendor_one_health_dept_license_3_code = "10" Then grh_vendor_one_health_dept_license_3_info = "Supportive Housing"

				EMReadScreen grh_vendor_one_number_of_licesned_beds, 4, 11, 24
				EMReadScreen grh_vendor_one_total_GRH_agreement_beds, 4, 11, 69
				EMReadScreen grh_vendor_one_resident_disa_type_1_code, 2,  12, 24
				If grh_vendor_one_resident_disa_type_1_code = "__" Then grh_vendor_one_resident_disa_type_1_info = ""
				If grh_vendor_one_resident_disa_type_1_code = "01" Then grh_vendor_one_resident_disa_type_1_info = "Development Disabled"
				If grh_vendor_one_resident_disa_type_1_code = "02" Then grh_vendor_one_resident_disa_type_1_info = "Chemically Dependent"
				If grh_vendor_one_resident_disa_type_1_code = "03" Then grh_vendor_one_resident_disa_type_1_info = "Mentally Ill"
				If grh_vendor_one_resident_disa_type_1_code = "04" Then grh_vendor_one_resident_disa_type_1_info = "Physically Handicapped"
				If grh_vendor_one_resident_disa_type_1_code = "05" Then grh_vendor_one_resident_disa_type_1_info = "Elderly"
				If grh_vendor_one_resident_disa_type_1_code = "06" Then grh_vendor_one_resident_disa_type_1_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_one_resident_disa_type_1_code = "08" Then grh_vendor_one_resident_disa_type_1_info = "None of the Above"

				EMReadScreen grh_vendor_one_resident_disa_type_2_code, 2,  12, 24
				If grh_vendor_one_resident_disa_type_2_code = "__" Then grh_vendor_one_resident_disa_type_2_info = ""
				If grh_vendor_one_resident_disa_type_2_code = "01" Then grh_vendor_one_resident_disa_type_2_info = "Development Disabled"
				If grh_vendor_one_resident_disa_type_2_code = "02" Then grh_vendor_one_resident_disa_type_2_info = "Chemically Dependent"
				If grh_vendor_one_resident_disa_type_2_code = "03" Then grh_vendor_one_resident_disa_type_2_info = "Mentally Ill"
				If grh_vendor_one_resident_disa_type_2_code = "04" Then grh_vendor_one_resident_disa_type_2_info = "Physically Handicapped"
				If grh_vendor_one_resident_disa_type_2_code = "05" Then grh_vendor_one_resident_disa_type_2_info = "Elderly"
				If grh_vendor_one_resident_disa_type_2_code = "06" Then grh_vendor_one_resident_disa_type_2_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_one_resident_disa_type_2_code = "08" Then grh_vendor_one_resident_disa_type_2_info = "None of the Above"

				EMReadScreen grh_vendor_one_resident_disa_type_3_code, 2,  12, 24
				If grh_vendor_one_resident_disa_type_3_code = "__" Then grh_vendor_one_resident_disa_type_3_info = ""
				If grh_vendor_one_resident_disa_type_3_code = "01" Then grh_vendor_one_resident_disa_type_3_info = "Development Disabled"
				If grh_vendor_one_resident_disa_type_3_code = "02" Then grh_vendor_one_resident_disa_type_3_info = "Chemically Dependent"
				If grh_vendor_one_resident_disa_type_3_code = "03" Then grh_vendor_one_resident_disa_type_3_info = "Mentally Ill"
				If grh_vendor_one_resident_disa_type_3_code = "04" Then grh_vendor_one_resident_disa_type_3_info = "Physically Handicapped"
				If grh_vendor_one_resident_disa_type_3_code = "05" Then grh_vendor_one_resident_disa_type_3_info = "Elderly"
				If grh_vendor_one_resident_disa_type_3_code = "06" Then grh_vendor_one_resident_disa_type_3_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_one_resident_disa_type_3_code = "08" Then grh_vendor_one_resident_disa_type_3_info = "None of the Above"

				EMReadScreen grh_vendor_one_resident_disa_type_4_code, 2,  12, 24
				If grh_vendor_one_resident_disa_type_4_code = "__" Then grh_vendor_one_resident_disa_type_4_info = ""
				If grh_vendor_one_resident_disa_type_4_code = "01" Then grh_vendor_one_resident_disa_type_4_info = "Development Disabled"
				If grh_vendor_one_resident_disa_type_4_code = "02" Then grh_vendor_one_resident_disa_type_4_info = "Chemically Dependent"
				If grh_vendor_one_resident_disa_type_4_code = "03" Then grh_vendor_one_resident_disa_type_4_info = "Mentally Ill"
				If grh_vendor_one_resident_disa_type_4_code = "04" Then grh_vendor_one_resident_disa_type_4_info = "Physically Handicapped"
				If grh_vendor_one_resident_disa_type_4_code = "05" Then grh_vendor_one_resident_disa_type_4_info = "Elderly"
				If grh_vendor_one_resident_disa_type_4_code = "06" Then grh_vendor_one_resident_disa_type_4_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_one_resident_disa_type_4_code = "08" Then grh_vendor_one_resident_disa_type_4_info = "None of the Above"

				EMReadScreen grh_vendor_one_resident_disa_type_5_code, 2,  12, 24
				If grh_vendor_one_resident_disa_type_5_code = "__" Then grh_vendor_one_resident_disa_type_5_info = ""
				If grh_vendor_one_resident_disa_type_5_code = "01" Then grh_vendor_one_resident_disa_type_5_info = "Development Disabled"
				If grh_vendor_one_resident_disa_type_5_code = "02" Then grh_vendor_one_resident_disa_type_5_info = "Chemically Dependent"
				If grh_vendor_one_resident_disa_type_5_code = "03" Then grh_vendor_one_resident_disa_type_5_info = "Mentally Ill"
				If grh_vendor_one_resident_disa_type_5_code = "04" Then grh_vendor_one_resident_disa_type_5_info = "Physically Handicapped"
				If grh_vendor_one_resident_disa_type_5_code = "05" Then grh_vendor_one_resident_disa_type_5_info = "Elderly"
				If grh_vendor_one_resident_disa_type_5_code = "06" Then grh_vendor_one_resident_disa_type_5_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_one_resident_disa_type_5_code = "08" Then grh_vendor_one_resident_disa_type_5_info = "None of the Above"

				EMReadScreen grh_vendor_one_room_and_board_rate_one_monthly, 8, 15, 54
				EMReadScreen grh_vendor_one_room_and_board_rate_one_per_diem, 8, 15, 68
				EMReadScreen grh_vendor_one_SSR_monthly, 8, 16, 54
				EMReadScreen grh_vendor_one_SSR_per_diem, 8, 16, 68

				grh_vendor_one_room_and_board_rate_one_monthly = replace(grh_vendor_one_room_and_board_rate_one_monthly, "_", "")
				grh_vendor_one_room_and_board_rate_one_per_diem = trim(grh_vendor_one_room_and_board_rate_one_per_diem)
				grh_vendor_one_SSR_monthly = replace(grh_vendor_one_SSR_monthly, "_", "")
				grh_vendor_one_SSR_per_diem = trim(grh_vendor_one_SSR_per_diem)

				PF3
			End If


			If grh_elig_budg_vendor_number_two <> "" Then
				Call write_value_and_transmit(grh_elig_budg_vendor_number_two, 4, 59)
				EMReadScreen grh_vendor_two_name, 					30, 3, 15
				EMReadScreen grh_vendor_two_c_o, 					30, 4, 15
				EMReadScreen grh_vendor_two_street_one, 			22, 5, 15
				EMReadScreen grh_vendor_two_street_two, 			22, 6, 15
				EMReadScreen grh_vendor_two_city, 					15, 7, 15
				EMReadScreen grh_vendor_two_state, 					2, 7, 36
				EMReadScreen grh_vendor_two_zip, 					10, 7, 46
				EMReadScreen grh_vendor_two_grh_yn, 				1, 4, 57
				EMReadScreen grh_vendor_two_non_profit_yn, 			1, 4, 78
				EMReadScreen grh_vendor_two_phone, 					16, 6, 54
				grh_vendor_two_phone = "(" & replace(replace(grh_vendor_two_phone, " )  ", ")"), "  ", "-")
				EMReadScreen grh_vendor_two_county, 				2, 7, 61
				EMReadScreen grh_vendor_two_status_code, 			1, 16, 15
				If grh_vendor_two_status_code = "A" Then grh_vendor_two_status_info = "Active"
				If grh_vendor_two_status_code = "D" Then grh_vendor_two_status_info = "Delete"
				If grh_vendor_two_status_code = "M" Then grh_vendor_two_status_info = "Merged"
				If grh_vendor_two_status_code = "P" Then grh_vendor_two_status_info = "Pending"
				If grh_vendor_two_status_code = "T" Then grh_vendor_two_status_info = "Terminated"
				EMReadScreen grh_vendor_two_incorporated_yn, 		1, 9, 22
				EMReadScreen grh_vendor_two_federal_tax_id, 			9, 9, 41
				EMReadScreen grh_vendor_two_ssn, 					11, 9, 61
				If grh_vendor_two_ssn = "___ __ ____" Then grh_vendor_two_ssn = ""
				grh_vendor_two_ssn = replace(grh_vendor_two_ssn, " ", "-")
				EMReadScreen grh_vendor_two_2nd_address_type_code, 	1, 10, 22
				If grh_vendor_two_2nd_address_type_code = "1" Then grh_vendor_two_2nd_address_type_info = "Mailing Address"
				If grh_vendor_two_2nd_address_type_code = "2" Then grh_vendor_two_2nd_address_type_info = "Court Order"
				EMReadScreen grh_vendor_two_2nd_address_eff_date, 	8, 11, 15
				If grh_vendor_two_2nd_address_eff_date = "__ __ __" Then grh_vendor_two_2nd_address_eff_date = ""
				grh_vendor_two_2nd_address_eff_date = replace(grh_vendor_two_2nd_address_eff_date, " ", "/")
				EMReadScreen grh_vendor_two_2nd_name, 				30, 11, 15
				EMReadScreen grh_vendor_two_2nd_c_o, 				30, 12, 15
				EMReadScreen grh_vendor_two_2nd_street_one, 		22, 13, 15
				EMReadScreen grh_vendor_two_2nd_street_two, 		22, 14, 15
				EMReadScreen grh_vendor_two_2nd_city, 				15, 15, 15
				EMReadScreen grh_vendor_two_2nd_state, 				2, 15, 35
				EMReadScreen grh_vendor_two_2nd_zip, 				10, 15, 44
				EMReadScreen grh_vendor_two_direct_deposit_yn, 		1, 12, 76
				EMReadScreen grh_vendor_two_merge_vendor_number, 	8, 16, 38
				EMReadScreen grh_vendor_two_acct_number_required_yn, 1, 17, 74
				EMReadScreen grh_vendor_two_blocked_county_numbers_list, 29, 18, 23

				grh_vendor_two_name = replace(grh_vendor_two_name, "_", "")
				grh_vendor_two_c_o = replace(grh_vendor_two_c_o, "_", "")
				grh_vendor_two_street_one = replace(grh_vendor_two_street_one, "_", "")
				grh_vendor_two_street_two = replace(grh_vendor_two_street_two, "_", "")
				grh_vendor_two_city = replace(grh_vendor_two_city, "_", "")
				grh_vendor_two_zip = trim(grh_vendor_two_zip)
				grh_vendor_two_zip = replace(grh_vendor_two_zip, " ", "-")

				grh_vendor_two_federal_tax_id = replace(grh_vendor_two_federal_tax_id, "_", "")

				grh_vendor_two_2nd_name = replace(grh_vendor_two_2nd_name, "_", "")
				grh_vendor_two_2nd_c_o = replace(grh_vendor_two_2nd_c_o, "_", "")
				grh_vendor_two_2nd_street_one = replace(grh_vendor_two_2nd_street_one, "_", "")
				grh_vendor_two_2nd_street_two = replace(grh_vendor_two_2nd_street_two, "_", "")
				grh_vendor_two_2nd_city = replace(grh_vendor_two_2nd_city, "_", "")
				grh_vendor_two_2nd_zip = replace(grh_vendor_two_2nd_zip, "_", "")
				grh_vendor_two_2nd_zip = trim(grh_vendor_two_2nd_zip)
				grh_vendor_two_2nd_zip = replace(grh_vendor_two_2nd_zip, " ", "-")

				grh_vendor_two_merge_vendor_number = replace(grh_vendor_two_merge_vendor_number, "_", "")
				grh_vendor_two_acct_number_required_yn = replace(grh_vendor_two_acct_number_required_yn, "_", "")

				grh_vendor_two_blocked_county_numbers_list = replace(grh_vendor_two_blocked_county_numbers_list, "_", "")
				grh_vendor_two_blocked_county_numbers_list = trim(grh_vendor_two_blocked_county_numbers_list)

				transmit
				EMReadScreen grh_vendor_two_current_rate_period_code, 1, 6, 24

				If grh_vendor_two_current_rate_period_code = "1" Then grh_vendor_two_current_rate_period_info = "Calendar Year"
				If grh_vendor_two_current_rate_period_code = "2" Then grh_vendor_two_current_rate_period_info = "Fiscal Year"
				If grh_vendor_two_current_rate_period_code = "3" Then grh_vendor_two_current_rate_period_info = "Federal Fiscal Year"
				If grh_vendor_two_current_rate_period_code = "4" Then grh_vendor_two_current_rate_period_info = "Other"

				EMReadScreen grh_vendor_two_rate_from_date, 7, 6, 47
				EMReadScreen grh_vendor_two_rate_to_date, 7, 6, 61
				EMReadScreen grh_vendor_two_initial_rate_date, 10, 7, 21
				EMReadScreen grh_vendor_two_NPI_number, 10, 7, 41
				EMReadScreen grh_vendor_two_family_foster_care_yn, 1, 8, 69
				EMReadScreen grh_vendor_two_rate_limit, 8, 9, 24
				EMReadScreen grh_vendor_two_exempt_reason_code, 2, 9, 69
				If grh_vendor_two_exempt_reason_code = "01" Then grh_vendor_two_exempt_reason_info = "Andrew Board & Care"
				If grh_vendor_two_exempt_reason_code = "04" Then grh_vendor_two_exempt_reason_info = "Aldrich"
				If grh_vendor_two_exempt_reason_code = "05" Then grh_vendor_two_exempt_reason_info = "Amy Johnson"
				If grh_vendor_two_exempt_reason_code = "09" Then grh_vendor_two_exempt_reason_info = "Quinlan Home"
				If grh_vendor_two_exempt_reason_code = "10" Then grh_vendor_two_exempt_reason_info = "Revere Home"
				If grh_vendor_two_exempt_reason_code = "11" Then grh_vendor_two_exempt_reason_info = "River Oaks"
				If grh_vendor_two_exempt_reason_code = "12" Then grh_vendor_two_exempt_reason_info = "Special Srvices"
				If grh_vendor_two_exempt_reason_code = "14" Then grh_vendor_two_exempt_reason_info = "Albert Lea"
				If grh_vendor_two_exempt_reason_code = "15" Then grh_vendor_two_exempt_reason_info = "Metro Demo"
				If grh_vendor_two_exempt_reason_code = "16" Then grh_vendor_two_exempt_reason_info = "Broadway"
				If grh_vendor_two_exempt_reason_code = "18" Then grh_vendor_two_exempt_reason_info = "Murphy's Board and Care"
				If grh_vendor_two_exempt_reason_code = "25" Then grh_vendor_two_exempt_reason_info = "Valley Home"
				If grh_vendor_two_exempt_reason_code = "26" Then grh_vendor_two_exempt_reason_info = "LTH Supportive Housing"
				If grh_vendor_two_exempt_reason_code = "27" Then grh_vendor_two_exempt_reason_info = "Boarding Care Home"
				If grh_vendor_two_exempt_reason_code = "28" Then grh_vendor_two_exempt_reason_info = "Banked Bed"
				If grh_vendor_two_exempt_reason_code = "29" Then grh_vendor_two_exempt_reason_info = "Tribe Certified Housing"

				EMReadScreen grh_vendor_two_DHS_license_1_code, 2, 10, 24
				If grh_vendor_two_DHS_license_1_code = "__" Then grh_vendor_two_DHS_license_1_info = ""
				If grh_vendor_two_DHS_license_1_code = "01" Then grh_vendor_two_DHS_license_1_info = "SILS- Developmental Disabled Rule 18"
				If grh_vendor_two_DHS_license_1_code = "02" Then grh_vendor_two_DHS_license_1_info = "Developmentaly Diabled Rule 34"
				If grh_vendor_two_DHS_license_1_code = "03" Then grh_vendor_two_DHS_license_1_info = "Adult Mentally Ill Rule 36"
				If grh_vendor_two_DHS_license_1_code = "04" Then grh_vendor_two_DHS_license_1_info = "Adult Foster Care Rule 203"
				If grh_vendor_two_DHS_license_1_code = "05" Then grh_vendor_two_DHS_license_1_info = "Mentally Retarded Waiver Rule 42"
				If grh_vendor_two_DHS_license_1_code = "06" Then grh_vendor_two_DHS_license_1_info = "Pregnant Woman Shelter Rule 6"
				If grh_vendor_two_DHS_license_1_code = "07" Then grh_vendor_two_DHS_license_1_info = "Other DHS license"
				If grh_vendor_two_DHS_license_1_code = "08" Then grh_vendor_two_DHS_license_1_info = "No DHS License"
				If grh_vendor_two_DHS_license_1_code = "09" Then grh_vendor_two_DHS_license_1_info = "Physical Handicap Rule 80"
				If grh_vendor_two_DHS_license_1_code = "10" Then grh_vendor_two_DHS_license_1_info = "Child Foster Care Rules 1 & 8"
				If grh_vendor_two_DHS_license_1_code = "11" Then grh_vendor_two_DHS_license_1_info = "Chemical Dependancy Rule 35"
				EMReadScreen grh_vendor_two_DHS_license_2_code, 2, 10, 27
				If grh_vendor_two_DHS_license_2_code = "__" Then grh_vendor_two_DHS_license_2_info = ""
				If grh_vendor_two_DHS_license_2_code = "01" Then grh_vendor_two_DHS_license_2_info = "SILS- Developmental Disabled Rule 18"
				If grh_vendor_two_DHS_license_2_code = "02" Then grh_vendor_two_DHS_license_2_info = "Developmentaly Diabled Rule 34"
				If grh_vendor_two_DHS_license_2_code = "03" Then grh_vendor_two_DHS_license_2_info = "Adult Mentally Ill Rule 36"
				If grh_vendor_two_DHS_license_2_code = "04" Then grh_vendor_two_DHS_license_2_info = "Adult Foster Care Rule 203"
				If grh_vendor_two_DHS_license_2_code = "05" Then grh_vendor_two_DHS_license_2_info = "Mentally Retarded Waiver Rule 42"
				If grh_vendor_two_DHS_license_2_code = "06" Then grh_vendor_two_DHS_license_2_info = "Pregnant Woman Shelter Rule 6"
				If grh_vendor_two_DHS_license_2_code = "07" Then grh_vendor_two_DHS_license_2_info = "Other DHS license"
				If grh_vendor_two_DHS_license_2_code = "08" Then grh_vendor_two_DHS_license_2_info = "No DHS License"
				If grh_vendor_two_DHS_license_2_code = "09" Then grh_vendor_two_DHS_license_2_info = "Physical Handicap Rule 80"
				If grh_vendor_two_DHS_license_2_code = "10" Then grh_vendor_two_DHS_license_2_info = "Child Foster Care Rules 1 & 8"
				If grh_vendor_two_DHS_license_2_code = "11" Then grh_vendor_two_DHS_license_2_info = "Chemical Dependancy Rule 35"
				EMReadScreen grh_vendor_two_DHS_license_3_code, 2, 10, 30
				If grh_vendor_two_DHS_license_3_code = "__" Then grh_vendor_two_DHS_license_3_info = ""
				If grh_vendor_two_DHS_license_3_code = "01" Then grh_vendor_two_DHS_license_3_info = "SILS- Developmental Disabled Rule 18"
				If grh_vendor_two_DHS_license_3_code = "02" Then grh_vendor_two_DHS_license_3_info = "Developmentaly Diabled Rule 34"
				If grh_vendor_two_DHS_license_3_code = "03" Then grh_vendor_two_DHS_license_3_info = "Adult Mentally Ill Rule 36"
				If grh_vendor_two_DHS_license_3_code = "04" Then grh_vendor_two_DHS_license_3_info = "Adult Foster Care Rule 203"
				If grh_vendor_two_DHS_license_3_code = "05" Then grh_vendor_two_DHS_license_3_info = "Mentally Retarded Waiver Rule 42"
				If grh_vendor_two_DHS_license_3_code = "06" Then grh_vendor_two_DHS_license_3_info = "Pregnant Woman Shelter Rule 6"
				If grh_vendor_two_DHS_license_3_code = "07" Then grh_vendor_two_DHS_license_3_info = "Other DHS license"
				If grh_vendor_two_DHS_license_3_code = "08" Then grh_vendor_two_DHS_license_3_info = "No DHS License"
				If grh_vendor_two_DHS_license_3_code = "09" Then grh_vendor_two_DHS_license_3_info = "Physical Handicap Rule 80"
				If grh_vendor_two_DHS_license_3_code = "10" Then grh_vendor_two_DHS_license_3_info = "Child Foster Care Rules 1 & 8"
				If grh_vendor_two_DHS_license_3_code = "11" Then grh_vendor_two_DHS_license_3_info = "Chemical Dependancy Rule 35"

				EMReadScreen grh_vendor_two_health_dept_license_1_code, 2, 10, 69
				If grh_vendor_two_health_dept_license_1_code = "__" Then grh_vendor_two_health_dept_license_1_info = ""
				If grh_vendor_two_health_dept_license_1_code = "01" Then grh_vendor_two_health_dept_license_1_info = "Nursing Home"
				If grh_vendor_two_health_dept_license_1_code = "02" Then grh_vendor_two_health_dept_license_1_info = "Boarding Care Home"
				If grh_vendor_two_health_dept_license_1_code = "03" Then grh_vendor_two_health_dept_license_1_info = "Supervised Living Facility"
				If grh_vendor_two_health_dept_license_1_code = "04" Then grh_vendor_two_health_dept_license_1_info = "Board and Lodging"
				If grh_vendor_two_health_dept_license_1_code = "05" Then grh_vendor_two_health_dept_license_1_info = "Hotal/Restaurant"
				If grh_vendor_two_health_dept_license_1_code = "06" Then grh_vendor_two_health_dept_license_1_info = "Board & Lodge with Special Services"
				If grh_vendor_two_health_dept_license_1_code = "07" Then grh_vendor_two_health_dept_license_1_info = "Tribal License"
				If grh_vendor_two_health_dept_license_1_code = "08" Then grh_vendor_two_health_dept_license_1_info = "Metro Demo"
				If grh_vendor_two_health_dept_license_1_code = "09" Then grh_vendor_two_health_dept_license_1_info = "Housing with Services"
				If grh_vendor_two_health_dept_license_1_code = "10" Then grh_vendor_two_health_dept_license_1_info = "Supportive Housing"
				EMReadScreen grh_vendor_two_health_dept_license_2_code, 2, 10, 72
				If grh_vendor_two_health_dept_license_2_code = "__" Then grh_vendor_two_health_dept_license_2_info = ""
				If grh_vendor_two_health_dept_license_2_code = "01" Then grh_vendor_two_health_dept_license_2_info = "Nursing Home"
				If grh_vendor_two_health_dept_license_2_code = "02" Then grh_vendor_two_health_dept_license_2_info = "Boarding Care Home"
				If grh_vendor_two_health_dept_license_2_code = "03" Then grh_vendor_two_health_dept_license_2_info = "Supervised Living Facility"
				If grh_vendor_two_health_dept_license_2_code = "04" Then grh_vendor_two_health_dept_license_2_info = "Board and Lodging"
				If grh_vendor_two_health_dept_license_2_code = "05" Then grh_vendor_two_health_dept_license_2_info = "Hotal/Restaurant"
				If grh_vendor_two_health_dept_license_2_code = "06" Then grh_vendor_two_health_dept_license_2_info = "Board & Lodge with Special Services"
				If grh_vendor_two_health_dept_license_2_code = "07" Then grh_vendor_two_health_dept_license_2_info = "Tribal License"
				If grh_vendor_two_health_dept_license_2_code = "08" Then grh_vendor_two_health_dept_license_2_info = "Metro Demo"
				If grh_vendor_two_health_dept_license_2_code = "09" Then grh_vendor_two_health_dept_license_2_info = "Housing with Services"
				If grh_vendor_two_health_dept_license_2_code = "10" Then grh_vendor_two_health_dept_license_2_info = "Supportive Housing"
				EMReadScreen grh_vendor_two_health_dept_license_3_code, 2, 10, 75
				If grh_vendor_two_health_dept_license_3_code = "__" Then grh_vendor_two_health_dept_license_3_info = ""
				If grh_vendor_two_health_dept_license_3_code = "01" Then grh_vendor_two_health_dept_license_3_info = "Nursing Home"
				If grh_vendor_two_health_dept_license_3_code = "02" Then grh_vendor_two_health_dept_license_3_info = "Boarding Care Home"
				If grh_vendor_two_health_dept_license_3_code = "03" Then grh_vendor_two_health_dept_license_3_info = "Supervised Living Facility"
				If grh_vendor_two_health_dept_license_3_code = "04" Then grh_vendor_two_health_dept_license_3_info = "Board and Lodging"
				If grh_vendor_two_health_dept_license_3_code = "05" Then grh_vendor_two_health_dept_license_3_info = "Hotal/Restaurant"
				If grh_vendor_two_health_dept_license_3_code = "06" Then grh_vendor_two_health_dept_license_3_info = "Board & Lodge with Special Services"
				If grh_vendor_two_health_dept_license_3_code = "07" Then grh_vendor_two_health_dept_license_3_info = "Tribal License"
				If grh_vendor_two_health_dept_license_3_code = "08" Then grh_vendor_two_health_dept_license_3_info = "Metro Demo"
				If grh_vendor_two_health_dept_license_3_code = "09" Then grh_vendor_two_health_dept_license_3_info = "Housing with Services"
				If grh_vendor_two_health_dept_license_3_code = "10" Then grh_vendor_two_health_dept_license_3_info = "Supportive Housing"

				EMReadScreen grh_vendor_two_number_of_licesned_beds, 4, 11, 24
				EMReadScreen grh_vendor_two_total_GRH_agreement_beds, 4, 11, 69
				EMReadScreen grh_vendor_two_resident_disa_type_1_code, 2,  12, 24
				If grh_vendor_two_resident_disa_type_1_code = "__" Then grh_vendor_two_resident_disa_type_1_info = ""
				If grh_vendor_two_resident_disa_type_1_code = "01" Then grh_vendor_two_resident_disa_type_1_info = "Development Disabled"
				If grh_vendor_two_resident_disa_type_1_code = "02" Then grh_vendor_two_resident_disa_type_1_info = "Chemically Dependent"
				If grh_vendor_two_resident_disa_type_1_code = "03" Then grh_vendor_two_resident_disa_type_1_info = "Mentally Ill"
				If grh_vendor_two_resident_disa_type_1_code = "04" Then grh_vendor_two_resident_disa_type_1_info = "Physically Handicapped"
				If grh_vendor_two_resident_disa_type_1_code = "05" Then grh_vendor_two_resident_disa_type_1_info = "Elderly"
				If grh_vendor_two_resident_disa_type_1_code = "06" Then grh_vendor_two_resident_disa_type_1_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_two_resident_disa_type_1_code = "08" Then grh_vendor_two_resident_disa_type_1_info = "None of the Above"

				EMReadScreen grh_vendor_two_resident_disa_type_2_code, 2,  12, 24
				If grh_vendor_two_resident_disa_type_2_code = "__" Then grh_vendor_two_resident_disa_type_2_info = ""
				If grh_vendor_two_resident_disa_type_2_code = "01" Then grh_vendor_two_resident_disa_type_2_info = "Development Disabled"
				If grh_vendor_two_resident_disa_type_2_code = "02" Then grh_vendor_two_resident_disa_type_2_info = "Chemically Dependent"
				If grh_vendor_two_resident_disa_type_2_code = "03" Then grh_vendor_two_resident_disa_type_2_info = "Mentally Ill"
				If grh_vendor_two_resident_disa_type_2_code = "04" Then grh_vendor_two_resident_disa_type_2_info = "Physically Handicapped"
				If grh_vendor_two_resident_disa_type_2_code = "05" Then grh_vendor_two_resident_disa_type_2_info = "Elderly"
				If grh_vendor_two_resident_disa_type_2_code = "06" Then grh_vendor_two_resident_disa_type_2_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_two_resident_disa_type_2_code = "08" Then grh_vendor_two_resident_disa_type_2_info = "None of the Above"

				EMReadScreen grh_vendor_two_resident_disa_type_3_code, 2,  12, 24
				If grh_vendor_two_resident_disa_type_3_code = "__" Then grh_vendor_two_resident_disa_type_3_info = ""
				If grh_vendor_two_resident_disa_type_3_code = "01" Then grh_vendor_two_resident_disa_type_3_info = "Development Disabled"
				If grh_vendor_two_resident_disa_type_3_code = "02" Then grh_vendor_two_resident_disa_type_3_info = "Chemically Dependent"
				If grh_vendor_two_resident_disa_type_3_code = "03" Then grh_vendor_two_resident_disa_type_3_info = "Mentally Ill"
				If grh_vendor_two_resident_disa_type_3_code = "04" Then grh_vendor_two_resident_disa_type_3_info = "Physically Handicapped"
				If grh_vendor_two_resident_disa_type_3_code = "05" Then grh_vendor_two_resident_disa_type_3_info = "Elderly"
				If grh_vendor_two_resident_disa_type_3_code = "06" Then grh_vendor_two_resident_disa_type_3_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_two_resident_disa_type_3_code = "08" Then grh_vendor_two_resident_disa_type_3_info = "None of the Above"

				EMReadScreen grh_vendor_two_resident_disa_type_4_code, 2,  12, 24
				If grh_vendor_two_resident_disa_type_4_code = "__" Then grh_vendor_two_resident_disa_type_4_info = ""
				If grh_vendor_two_resident_disa_type_4_code = "01" Then grh_vendor_two_resident_disa_type_4_info = "Development Disabled"
				If grh_vendor_two_resident_disa_type_4_code = "02" Then grh_vendor_two_resident_disa_type_4_info = "Chemically Dependent"
				If grh_vendor_two_resident_disa_type_4_code = "03" Then grh_vendor_two_resident_disa_type_4_info = "Mentally Ill"
				If grh_vendor_two_resident_disa_type_4_code = "04" Then grh_vendor_two_resident_disa_type_4_info = "Physically Handicapped"
				If grh_vendor_two_resident_disa_type_4_code = "05" Then grh_vendor_two_resident_disa_type_4_info = "Elderly"
				If grh_vendor_two_resident_disa_type_4_code = "06" Then grh_vendor_two_resident_disa_type_4_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_two_resident_disa_type_4_code = "08" Then grh_vendor_two_resident_disa_type_4_info = "None of the Above"

				EMReadScreen grh_vendor_two_resident_disa_type_5_code, 2,  12, 24
				If grh_vendor_two_resident_disa_type_5_code = "__" Then grh_vendor_two_resident_disa_type_5_info = ""
				If grh_vendor_two_resident_disa_type_5_code = "01" Then grh_vendor_two_resident_disa_type_5_info = "Development Disabled"
				If grh_vendor_two_resident_disa_type_5_code = "02" Then grh_vendor_two_resident_disa_type_5_info = "Chemically Dependent"
				If grh_vendor_two_resident_disa_type_5_code = "03" Then grh_vendor_two_resident_disa_type_5_info = "Mentally Ill"
				If grh_vendor_two_resident_disa_type_5_code = "04" Then grh_vendor_two_resident_disa_type_5_info = "Physically Handicapped"
				If grh_vendor_two_resident_disa_type_5_code = "05" Then grh_vendor_two_resident_disa_type_5_info = "Elderly"
				If grh_vendor_two_resident_disa_type_5_code = "06" Then grh_vendor_two_resident_disa_type_5_info = "Deaf/Blind or Brain Injured"
				If grh_vendor_two_resident_disa_type_5_code = "08" Then grh_vendor_two_resident_disa_type_5_info = "None of the Above"

				EMReadScreen grh_vendor_two_room_and_board_rate_two_monthly, 8, 15, 54
				EMReadScreen grh_vendor_two_room_and_board_rate_two_per_diem, 8, 15, 68
				EMReadScreen grh_vendor_two_SSR_monthly, 8, 16, 54
				EMReadScreen grh_vendor_two_SSR_per_diem, 8, 16, 68

				grh_vendor_two_room_and_board_rate_one_monthly = replace(grh_vendor_two_room_and_board_rate_one_monthly, "_", "")
				grh_vendor_two_room_and_board_rate_one_per_diem = trim(grh_vendor_two_room_and_board_rate_one_per_diem)
				grh_vendor_two_SSR_monthly = replace(grh_vendor_two_SSR_monthly, "_", "")
				grh_vendor_two_SSR_per_diem = trim(grh_vendor_two_SSR_per_diem)

				PF3
			End If
		End if

		Call back_to_SELF
	end sub

end class

class emer_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found

	public initial_search_month
	public initial_search_year

	public emer_program

	public emer_check_issue_date()
	public emer_check_program()
	public emer_check_status_code()
	public emer_check_status_info()
	public emer_check_warrant_number()
	public emer_check_transaction_amount()
	public emer_check_type_code()
	public emer_check_type_info()
	public emer_check_transaction_number()
	public emer_check_from_date()
	public emer_check_to_date()
	public emer_check_payment_reason()
	public emer_check_payment_to_name()
	public emer_check_payment_to_address()
	public emer_check_mail_method()
	public emer_check_payment_method()
	public emer_check_vendor_number()
	public emer_check_fiche_number()
	public emer_check_payment_amount()
	public emer_check_entitement_amount()
	public emer_check_recoupment_amount()
	public emer_check_replacement_amount()
	public emer_check_cacnel_amount()
	public emer_check_food_portion_amount()
	public emer_check_reconciliation_date()
	public emer_check_cancel_reason()
	public emer_check_replacement_reason()
	public emer_check_picup_status()
	public emer_check_pickup_date()
	public emer_check_servicing_county()
	public emer_check_responsibility_county()
	public emer_check_adjusting_transaction()
	public emer_check_original_transaction()
	public emer_check_vendor_name()
	public emer_check_vendor_c_o()
	public emer_check_vendor_street_one()
	public emer_check_vendor_street_two()
	public emer_check_vendor_city()
	public emer_check_vendor_state()
	public emer_check_vendor_zip()
	public emer_check_vendor_grh_yn()
	public emer_check_vendor_non_profit_yn()
	public emer_check_vendor_phone()
	public emer_check_vendor_county()
	public emer_check_vendor_status_code()
	public emer_check_vendor_status_info()
	public emer_check_vendor_incorporated_yn()
	public emer_check_vendor_federal_tax_id()
	public emer_check_vendor_ssn()
	public emer_check_vendor_2nd_address_type_code()
	public emer_check_vendor_2nd_address_type_info()
	public emer_check_vendor_2nd_address_eff_date()
	public emer_check_vendor_2nd_name()
	public emer_check_vendor_2nd_c_o()
	public emer_check_vendor_2nd_street_one()
	public emer_check_vendor_2nd_street_two()
	public emer_check_vendor_2nd_city()
	public emer_check_vendor_2nd_state()
	public emer_check_vendor_2nd_zip()
	public emer_check_vendor_direct_deposit_yn()
	public emer_check_vendor_merge_vendor_number()
	public emer_check_vendor_acct_number_required_yn()
	public emer_check_vendor_blocked_county_numbers_list()

	public emer_elig_case_test_citizenship
	public emer_elig_case_test_coop_MFIP
	public emer_elig_case_test_copayment
	public emer_elig_case_test_cost_effective
	public emer_elig_case_test_eligible_child
	public emer_elig_case_test_emergency
	public emer_elig_case_test_equitable_interest
	public emer_elig_case_test_residency
	public emer_elig_case_test_resources
	public emer_elig_case_test_verif
	public emer_elig_case_test_12_month
	public emer_elig_case_test_coop_work
	public emer_elig_case_test_county_allocation
	public emer_elig_case_test_elig_other_program
	public emer_elig_case_test_200_percent_fpg

	public emer_elig_available_gross_earned_income
	public emer_elig_available_actual_work_expense
	public emer_elig_available_net_earned_income
	public emer_elig_available_unearned_income
	public emer_elig_available_assets
	public emer_elig_available_other_assets
	public emer_elig_available_total_income_assets
	public emer_elig_expense_rent_mortgage
	public emer_elig_expense_fuel
	public emer_elig_expense_electric
	public emer_elig_expense_msa_standard
	public emer_elig_expense_car_payment
	public emer_elig_expense_phone
	public emer_elig_expense_food
	public emer_elig_expense_other
	public emer_elig_total_basic_needs
	public emer_elig_expense_net_income_assets

	public emer_elig_approved_date
	public emer_elig_process_date
	public emer_elig_summ_date_last_approval
	public emer_elig_summ_current_program_status
	public emer_elig_summ_eligibility_result
	public emer_elig_summ_last_used

	public emer_elig_summ_adults_in_unit
	public emer_elig_summ_children_in_unit
	public emer_elig_summ_begin_date
	public emer_elig_summ_end_date

	public emer_elig_summ_need_foreclosure
	public emer_elig_summ_need_temp_shelter
	public emer_elig_summ_need_other_shelter
	public emer_elig_summ_need_utility
	public emer_elig_summ_need_other
	public emer_elig_summ_need_total
	public emer_elig_summ_payment


	public emer_elig_ref_numbs()
	public emer_elig_membs_full_name()
	public emer_elig_membs_request_yn()
	public emer_elig_membs_code()
	public emer_elig_membs_info()
	public emer_elig_membs_fund_fact()
	public emer_elig_membs_adult_or_child()
	public emer_elig_membs_elig_status()
	public emer_elig_membs_12_month_test()
	public emer_elig_membs_last_emer_begin_date()

	public sub read_elig()
		approved_today = False
		approved_version_found = False

		ReDim emer_check_issue_date(0)
		ReDim emer_check_program(0)
		ReDim emer_check_status_code(0)
		ReDim emer_check_status_info(0)
		ReDim emer_check_warrant_number(0)
		ReDim emer_check_transaction_amount(0)
		ReDim emer_check_type_code(0)
		ReDim emer_check_type_info(0)
		ReDim emer_check_transaction_number(0)
		ReDim emer_check_from_date(0)
		ReDim emer_check_to_date(0)
		ReDim emer_check_payment_reason(0)
		ReDim emer_check_payment_to_name(0)
		ReDim emer_check_payment_to_address(0)
		ReDim emer_check_mail_method(0)
		ReDim emer_check_payment_method(0)
		ReDim emer_check_vendor_number(0)
		ReDim emer_check_fiche_number(0)
		ReDim emer_check_payment_amount(0)
		ReDim emer_check_entitement_amount(0)
		ReDim emer_check_recoupment_amount(0)
		ReDim emer_check_replacement_amount(0)
		ReDim emer_check_cacnel_amount(0)
		ReDim emer_check_food_portion_amount(0)
		ReDim emer_check_reconciliation_date(0)
		ReDim emer_check_cancel_reason(0)
		ReDim emer_check_replacement_reason(0)
		ReDim emer_check_picup_status(0)
		ReDim emer_check_pickup_date(0)
		ReDim emer_check_servicing_county(0)
		ReDim emer_check_responsibility_county(0)
		ReDim emer_check_adjusting_transaction(0)
		ReDim emer_check_original_transaction(0)
		ReDim emer_check_vendor_name(0)
		ReDim emer_check_vendor_c_o(0)
		ReDim emer_check_vendor_street_one(0)
		ReDim emer_check_vendor_street_two(0)
		ReDim emer_check_vendor_city(0)
		ReDim emer_check_vendor_state(0)
		ReDim emer_check_vendor_zip(0)
		ReDim emer_check_vendor_grh_yn(0)
		ReDim emer_check_vendor_non_profit_yn(0)
		ReDim emer_check_vendor_phone(0)
		ReDim emer_check_vendor_county(0)
		ReDim emer_check_vendor_status_code(0)
		ReDim emer_check_vendor_status_info(0)
		ReDim emer_check_vendor_incorporated_yn(0)
		ReDim emer_check_vendor_federal_tax_id(0)
		ReDim emer_check_vendor_ssn(0)
		ReDim emer_check_vendor_2nd_address_type_code(0)
		ReDim emer_check_vendor_2nd_address_type_info(0)
		ReDim emer_check_vendor_2nd_address_eff_date(0)
		ReDim emer_check_vendor_2nd_name(0)
		ReDim emer_check_vendor_2nd_c_o(0)
		ReDim emer_check_vendor_2nd_street_one(0)
		ReDim emer_check_vendor_2nd_street_two(0)
		ReDim emer_check_vendor_2nd_city(0)
		ReDim emer_check_vendor_2nd_state(0)
		ReDim emer_check_vendor_2nd_zip(0)
		ReDim emer_check_vendor_direct_deposit_yn(0)
		ReDim emer_check_vendor_merge_vendor_number(0)
		ReDim emer_check_vendor_acct_number_required_yn(0)
		ReDim emer_check_vendor_blocked_county_numbers_list(0)

		ReDim emer_elig_ref_numbs(0)
		ReDim emer_elig_membs_full_name(0)
		ReDim emer_elig_membs_request_yn(0)
		ReDim emer_elig_membs_code(0)
		ReDim emer_elig_membs_info(0)
		ReDim emer_elig_membs_fund_fact(0)
		ReDim emer_elig_membs_adult_or_child(0)
		ReDim emer_elig_membs_elig_status(0)
		ReDim emer_elig_membs_12_month_test(0)
		ReDim emer_elig_membs_last_emer_begin_date(0)

		Call navigate_to_MAXIS_screen("MONY", "INQX")
		EMWriteScreen initial_search_month, 6, 38
		EMWriteScreen initial_search_year, 6, 41
		EMWriteScreen CM_plus_1_mo, 6, 53
		EMWriteScreen CM_plus_1_yr, 6, 56
		EMWriteScreen "X", 9, 50
		EMWriteScreen "X", 11, 50
		EMWriteScreen "X", 12, 50
		transmit

		inqd_row = 6
		tx_count = 0
		EMReadScreen chck_prog, 7, inqd_row, 16
		chck_prog = trim(chck_prog)

		Do while chck_prog <> ""
			ReDim preserve emer_check_issue_date(tx_count)
			ReDim preserve emer_check_program(tx_count)
			ReDim preserve emer_check_status_code(tx_count)
			ReDim preserve emer_check_status_info(tx_count)
			ReDim preserve emer_check_warrant_number(tx_count)
			ReDim preserve emer_check_transaction_amount(tx_count)
			ReDim preserve emer_check_type_code(tx_count)
			ReDim preserve emer_check_type_info(tx_count)
			ReDim preserve emer_check_transaction_number(tx_count)
			ReDim preserve emer_check_from_date(tx_count)
			ReDim preserve emer_check_to_date(tx_count)
			ReDim preserve emer_check_payment_reason(tx_count)
			ReDim preserve emer_check_payment_to_name(tx_count)
			ReDim preserve emer_check_payment_to_address(tx_count)
			ReDim preserve emer_check_mail_method(tx_count)
			ReDim preserve emer_check_payment_method(tx_count)
			ReDim preserve emer_check_vendor_number(tx_count)
			ReDim preserve emer_check_fiche_number(tx_count)
			ReDim preserve emer_check_payment_amount(tx_count)
			ReDim preserve emer_check_entitement_amount(tx_count)
			ReDim preserve emer_check_recoupment_amount(tx_count)
			ReDim preserve emer_check_replacement_amount(tx_count)
			ReDim preserve emer_check_cacnel_amount(tx_count)
			ReDim preserve emer_check_food_portion_amount(tx_count)
			ReDim preserve emer_check_reconciliation_date(tx_count)
			ReDim preserve emer_check_cancel_reason(tx_count)
			ReDim preserve emer_check_replacement_reason(tx_count)
			ReDim preserve emer_check_picup_status(tx_count)
			ReDim preserve emer_check_pickup_date(tx_count)
			ReDim preserve emer_check_servicing_county(tx_count)
			ReDim preserve emer_check_responsibility_county(tx_count)
			ReDim preserve emer_check_adjusting_transaction(tx_count)
			ReDim preserve emer_check_original_transaction(tx_count)
			ReDim preserve emer_check_vendor_name(tx_count)
			ReDim preserve emer_check_vendor_c_o(tx_count)
			ReDim preserve emer_check_vendor_street_one(tx_count)
			ReDim preserve emer_check_vendor_street_two(tx_count)
			ReDim preserve emer_check_vendor_city(tx_count)
			ReDim preserve emer_check_vendor_state(tx_count)
			ReDim preserve emer_check_vendor_zip(tx_count)
			ReDim preserve emer_check_vendor_grh_yn(tx_count)
			ReDim preserve emer_check_vendor_non_profit_yn(tx_count)
			ReDim preserve emer_check_vendor_phone(tx_count)
			ReDim preserve emer_check_vendor_county(tx_count)
			ReDim preserve emer_check_vendor_status_code(tx_count)
			ReDim preserve emer_check_vendor_status_info(tx_count)
			ReDim preserve emer_check_vendor_incorporated_yn(tx_count)
			ReDim preserve emer_check_vendor_federal_tax_id(tx_count)
			ReDim preserve emer_check_vendor_ssn(tx_count)
			ReDim preserve emer_check_vendor_2nd_address_type_code(tx_count)
			ReDim preserve emer_check_vendor_2nd_address_type_info(tx_count)
			ReDim preserve emer_check_vendor_2nd_address_eff_date(tx_count)
			ReDim preserve emer_check_vendor_2nd_name(tx_count)
			ReDim preserve emer_check_vendor_2nd_c_o(tx_count)
			ReDim preserve emer_check_vendor_2nd_street_one(tx_count)
			ReDim preserve emer_check_vendor_2nd_street_two(tx_count)
			ReDim preserve emer_check_vendor_2nd_city(tx_count)
			ReDim preserve emer_check_vendor_2nd_state(tx_count)
			ReDim preserve emer_check_vendor_2nd_zip(tx_count)
			ReDim preserve emer_check_vendor_direct_deposit_yn(tx_count)
			ReDim preserve emer_check_vendor_merge_vendor_number(tx_count)
			ReDim preserve emer_check_vendor_acct_number_required_yn(tx_count)
			ReDim preserve emer_check_vendor_blocked_county_numbers_list(tx_count)

			emer_check_program(tx_count) = chck_prog
			EMReadScreen emer_check_issue_date(tx_count), 8, inqd_row, 7
			EMReadScreen emer_check_status_code(tx_count), 1, inqd_row, 26
			If emer_check_status_code(tx_count) = "C" Then emer_check_status_info(tx_count) = "Cancel/Return"
			If emer_check_status_code(tx_count) = "D" Then emer_check_status_info(tx_count) = "Denied"
			If emer_check_status_code(tx_count) = "I" Then emer_check_status_info(tx_count) = "Issued"
			If emer_check_status_code(tx_count) = "P" Then emer_check_status_info(tx_count) = "Pending"
			If emer_check_status_code(tx_count) = "R" Then emer_check_status_info(tx_count) = "Cashed"
			If emer_check_status_code(tx_count) = "S" Then emer_check_status_info(tx_count) = "Partial Cancel"
			If emer_check_status_code(tx_count) = "T" Then emer_check_status_info(tx_count) = "Stopped/Cashed"
			If emer_check_status_code(tx_count) = "X" Then emer_check_status_info(tx_count) = "Stopped"
			If emer_check_status_code(tx_count) = "B" Then emer_check_status_info(tx_count) = "Cashed and Replaced"
			EMReadScreen emer_check_warrant_number(tx_count), 8, inqd_row, 28
			EMReadScreen emer_check_transaction_amount(tx_count), 9, inqd_row, 37
			emer_check_transaction_amount(tx_count) = trim(emer_check_transaction_amount(tx_count))
			EMReadScreen emer_check_type_code(tx_count), 1, inqd_row, 48
			If emer_check_type_code(tx_count) = "1" Then emer_check_type_info(tx_count) = "Ongoing Issuance"
			If emer_check_type_code(tx_count) = "2" Then emer_check_type_info(tx_count) = "Same Day Local Issuance"
			If emer_check_type_code(tx_count) = "3" Then emer_check_type_info(tx_count) = "Replacement Issuance"
			If emer_check_type_code(tx_count) = "4" Then emer_check_type_info(tx_count) = "Same Day Issuance"
			If emer_check_type_code(tx_count) = "5" Then emer_check_type_info(tx_count) = "Nightly Issuance"
			If emer_check_type_code(tx_count) = "6" Then emer_check_type_info(tx_count) = "Manual Issuance"
			If emer_check_type_code(tx_count) = "7" Then emer_check_type_info(tx_count) = "EBT Rapid Electronic Issuance"
			If emer_check_type_code(tx_count) = "8" Then emer_check_type_info(tx_count) = "EBT Rapid Electronic Replacement"
			EMReadScreen emer_check_transaction_number(tx_count), 9, inqd_row, 51
			EMReadScreen emer_check_from_date(tx_count), 8, inqd_row, 62
			EMReadScreen emer_check_to_date(tx_count), 8, inqd_row, 73

			Call write_value_and_transmit("I", inqd_row, 4)


			EMReadScreen emer_check_payment_reason(tx_count), 	30, 7, 17
			EMReadScreen emer_check_payment_to_name(tx_count), 	30, 8, 17
			EMReadScreen addr_one, 								30, 9, 17
			EMReadScreen addr_two, 								30, 10, 17
			emer_check_payment_to_address(tx_count) = trim(trim(addr_one) & " " & trim(addr_two))
			EMReadScreen emer_check_mail_method(tx_count), 			15, 4, 63
			EMReadScreen emer_check_payment_method(tx_count), 		15, 5, 63
			EMReadScreen emer_check_vendor_number(tx_count), 		15, 6, 63
			EMReadScreen emer_check_fiche_number(tx_count), 		15, 7, 63
			EMReadScreen emer_check_payment_amount(tx_count), 		10, 13, 16
			EMReadScreen emer_check_entitement_amount(tx_count), 	10, 14, 16
			EMReadScreen emer_check_recoupment_amount(tx_count), 	10, 15, 16
			EMReadScreen emer_check_replacement_amount(tx_count), 	10, 16, 16
			EMReadScreen emer_check_cacnel_amount(tx_count), 		10, 17, 16
			EMReadScreen emer_check_food_portion_amount(tx_count), 	10, 18, 16
			EMReadScreen emer_check_reconciliation_date(tx_count), 	8, 6, 43
			EMReadScreen emer_check_cancel_reason(tx_count), 		30, 17, 41
			EMReadScreen emer_check_replacement_reason(tx_count), 	30, 18, 46
			EMReadScreen emer_check_picup_status(tx_count), 		10, 10, 70
			EMReadScreen emer_check_pickup_date(tx_count), 			8, 11, 70
			EMReadScreen emer_check_servicing_county(tx_count), 	2, 13, 70
			EMReadScreen emer_check_responsibility_county(tx_count), 2, 14, 70
			EMReadScreen emer_check_adjusting_transaction(tx_count), 10, 15, 70
			EMReadScreen emer_check_original_transaction(tx_count), 10, 16, 70

			emer_check_payment_reason(tx_count) = trim(emer_check_payment_reason(tx_count))
			emer_check_payment_to_name(tx_count) = trim(emer_check_payment_to_name(tx_count))
			emer_check_payment_to_address(tx_count) = trim(emer_check_payment_to_address(tx_count))
			emer_check_mail_method(tx_count) = trim(emer_check_mail_method(tx_count))
			emer_check_payment_method(tx_count) = trim(emer_check_payment_method(tx_count))
			emer_check_vendor_number(tx_count) = trim(emer_check_vendor_number(tx_count))
			emer_check_fiche_number(tx_count) = trim(emer_check_fiche_number(tx_count))
			emer_check_payment_amount(tx_count) = trim(emer_check_payment_amount(tx_count))
			emer_check_entitement_amount(tx_count) = trim(emer_check_entitement_amount(tx_count))
			emer_check_recoupment_amount(tx_count) = trim(emer_check_recoupment_amount(tx_count))
			emer_check_replacement_amount(tx_count) = trim(emer_check_replacement_amount(tx_count))
			emer_check_cacnel_amount(tx_count) = trim(emer_check_cacnel_amount(tx_count))
			emer_check_food_portion_amount(tx_count) = trim(emer_check_food_portion_amount(tx_count))
			emer_check_reconciliation_date(tx_count) = trim(emer_check_reconciliation_date(tx_count))
			emer_check_cancel_reason(tx_count) = trim(emer_check_cancel_reason(tx_count))
			emer_check_replacement_reason(tx_count) = trim(emer_check_replacement_reason(tx_count))
			emer_check_picup_status(tx_count) = trim(emer_check_picup_status(tx_count))
			emer_check_pickup_date(tx_count) = trim(emer_check_pickup_date(tx_count))
			emer_check_servicing_county(tx_count) = trim(emer_check_servicing_county(tx_count))
			emer_check_responsibility_county(tx_count) = trim(emer_check_responsibility_county(tx_count))
			emer_check_adjusting_transaction(tx_count) = trim(emer_check_adjusting_transaction(tx_count))
			emer_check_original_transaction(tx_count) = trim(emer_check_original_transaction(tx_count))

			PF3

			tx_count = tx_count + 1
			inqd_row = inqd_row + 1
			EMReadScreen chck_prog, 7, inqd_row, 16
			chck_prog = trim(chck_prog)
		Loop
		PF3

		for each_trans = 0 to UBound(emer_check_program)
			Call navigate_to_MAXIS_screen("MONY", "VNDS")


			Call write_value_and_transmit(emer_check_vendor_number(each_trans), 4, 59)
			EMReadScreen emer_check_vendor_name(each_trans), 					30, 3, 15
			EMReadScreen emer_check_vendor_c_o(each_trans), 					30, 4, 15
			EMReadScreen emer_check_vendor_street_one(each_trans), 				22, 5, 15
			EMReadScreen emer_check_vendor_street_two(each_trans), 				22, 6, 15
			EMReadScreen emer_check_vendor_city(each_trans), 					15, 7, 15
			EMReadScreen emer_check_vendor_state(each_trans), 					2, 7, 36
			EMReadScreen emer_check_vendor_zip(each_trans), 					10, 7, 46
			EMReadScreen emer_check_vendor_grh_yn(each_trans), 					1, 4, 57
			EMReadScreen emer_check_vendor_non_profit_yn(each_trans), 			1, 4, 78
			EMReadScreen emer_check_vendor_phone(each_trans), 					16, 6, 54
			emer_check_vendor_phone(each_trans) = "(" & replace(replace(emer_check_vendor_phone(each_trans), " )  ", ")"), "  ", "-")
			EMReadScreen emer_check_vendor_county(each_trans), 					2, 7, 61
			EMReadScreen emer_check_vendor_status_code(each_trans), 			1, 16, 15
			If emer_check_vendor_status_code(each_trans) = "A" Then emer_check_vendor_status_info(each_trans) = "Active"
			If emer_check_vendor_status_code(each_trans) = "D" Then emer_check_vendor_status_info(each_trans) = "Delete"
			If emer_check_vendor_status_code(each_trans) = "M" Then emer_check_vendor_status_info(each_trans) = "Merged"
			If emer_check_vendor_status_code(each_trans) = "P" Then emer_check_vendor_status_info(each_trans) = "Pending"
			If emer_check_vendor_status_code(each_trans) = "T" Then emer_check_vendor_status_info(each_trans) = "Terminated"
			EMReadScreen emer_check_vendor_incorporated_yn(each_trans), 		1, 9, 22
			EMReadScreen emer_check_vendor_federal_tax_id(each_trans), 			9, 9, 41
			EMReadScreen emer_check_vendor_ssn(each_trans), 					11, 9, 61
			If emer_check_vendor_ssn(each_trans) = "___ __ ____" Then emer_check_vendor_ssn(each_trans) = ""
			emer_check_vendor_ssn(each_trans) = replace(emer_check_vendor_ssn(each_trans), " ", "-")
			EMReadScreen emer_check_vendor_2nd_address_type_code(each_trans), 	1, 10, 22
			If emer_check_vendor_2nd_address_type_code(each_trans) = "1" Then emer_check_vendor_2nd_address_type_info(each_trans) = "Mailing Address"
			If emer_check_vendor_2nd_address_type_code(each_trans) = "2" Then emer_check_vendor_2nd_address_type_info(each_trans) = "Court Order"
			EMReadScreen emer_check_vendor_2nd_address_eff_date(each_trans), 	8, 11, 15
			If emer_check_vendor_2nd_address_eff_date(each_trans) = "__ __ __" Then emer_check_vendor_2nd_address_eff_date(each_trans) = ""
			emer_check_vendor_2nd_address_eff_date(each_trans) = replace(emer_check_vendor_2nd_address_eff_date(each_trans), " ", "/")
			EMReadScreen emer_check_vendor_2nd_name(each_trans), 				30, 11, 15
			EMReadScreen emer_check_vendor_2nd_c_o(each_trans), 				30, 12, 15
			EMReadScreen emer_check_vendor_2nd_street_one(each_trans), 			22, 13, 15
			EMReadScreen emer_check_vendor_2nd_street_two(each_trans), 			22, 14, 15
			EMReadScreen emer_check_vendor_2nd_city(each_trans), 				15, 15, 15
			EMReadScreen emer_check_vendor_2nd_state(each_trans), 				2, 15, 35
			EMReadScreen emer_check_vendor_2nd_zip(each_trans), 				10, 15, 44
			EMReadScreen emer_check_vendor_direct_deposit_yn(each_trans), 		1, 12, 76
			EMReadScreen emer_check_vendor_merge_vendor_number(each_trans), 	8, 16, 38
			EMReadScreen emer_check_vendor_acct_number_required_yn(each_trans), 1, 17, 74
			EMReadScreen emer_check_vendor_blocked_county_numbers_list(each_trans), 29, 18, 23

			emer_check_vendor_name(each_trans) = replace(emer_check_vendor_name(each_trans), "_", "")
			emer_check_vendor_c_o(each_trans) = replace(emer_check_vendor_c_o(each_trans), "_", "")
			emer_check_vendor_street_one(each_trans) = replace(emer_check_vendor_street_one(each_trans), "_", "")
			emer_check_vendor_street_two(each_trans) = replace(emer_check_vendor_street_two(each_trans), "_", "")
			emer_check_vendor_city(each_trans) = replace(emer_check_vendor_city(each_trans), "_", "")
			emer_check_vendor_zip(each_trans) = trim(emer_check_vendor_zip(each_trans))
			emer_check_vendor_zip(each_trans) = replace(emer_check_vendor_zip(each_trans), " ", "-")

			emer_check_vendor_federal_tax_id(each_trans) = replace(emer_check_vendor_federal_tax_id(each_trans), "_", "")

			emer_check_vendor_2nd_name(each_trans) = replace(emer_check_vendor_2nd_name(each_trans), "_", "")
			emer_check_vendor_2nd_c_o(each_trans) = replace(emer_check_vendor_2nd_c_o(each_trans), "_", "")
			emer_check_vendor_2nd_street_one(each_trans) = replace(emer_check_vendor_2nd_street_one(each_trans), "_", "")
			emer_check_vendor_2nd_street_two(each_trans) = replace(emer_check_vendor_2nd_street_two(each_trans), "_", "")
			emer_check_vendor_2nd_city(each_trans) = replace(emer_check_vendor_2nd_city(each_trans), "_", "")
			emer_check_vendor_2nd_zip(each_trans) = replace(emer_check_vendor_2nd_zip(each_trans), "_", "")
			emer_check_vendor_2nd_zip(each_trans) = trim(emer_check_vendor_2nd_zip(each_trans))
			emer_check_vendor_2nd_zip(each_trans) = replace(emer_check_vendor_2nd_zip(each_trans), " ", "-")

			emer_check_vendor_merge_vendor_number(each_trans) = replace(emer_check_vendor_merge_vendor_number(each_trans), "_", "")
			emer_check_vendor_acct_number_required_yn(each_trans) = replace(emer_check_vendor_acct_number_required_yn(each_trans), "_", "")

			emer_check_vendor_blocked_county_numbers_list(each_trans) = replace(emer_check_vendor_blocked_county_numbers_list(each_trans), "_", "")
			emer_check_vendor_blocked_county_numbers_list(each_trans) = trim((emer_check_vendor_blocked_county_numbers_list(each_trans)))

			PF3
		Next

		call navigate_to_MAXIS_screen("ELIG", "    ")
		EMWriteScreen elig_footer_month, 20, 55
		EMWriteScreen elig_footer_year, 20, 58
		call navigate_to_MAXIS_screen("ELIG", "EMER")
		Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		If approved_version_found = True Then
			EMReadScreen emer_program, 2, 4, 45

			ff_col = 59
			ac_col = 64
			es_col = 69
			If emer_program = "EA" Then
				rn_col = 6
				fn_col = 9
				rq_col = 33
				mc_col = 38
			End If
			If emer_program = "EG" Then
				rn_col = 8
				fn_col = 13
				rq_col = 37
				mc_col = 44
			End If

			emer_row = 8
			memb_count = 0
			Do
				EMReadScreen ref_numb, 2, emer_row, rn_col

				ReDim preserve emer_elig_ref_numbs(memb_count)
				ReDim preserve emer_elig_membs_full_name(memb_count)
				ReDim preserve emer_elig_membs_request_yn(memb_count)
				ReDim preserve emer_elig_membs_code(memb_count)
				ReDim preserve emer_elig_membs_info(memb_count)
				ReDim preserve emer_elig_membs_fund_fact(memb_count)
				ReDim preserve emer_elig_membs_adult_or_child(memb_count)
				ReDim preserve emer_elig_membs_elig_status(memb_count)
				ReDim preserve emer_elig_membs_12_month_test(memb_count)
				ReDim preserve emer_elig_membs_last_emer_begin_date(memb_count)

				emer_elig_ref_numbs(memb_count) = ref_numb
				EMReadScreen emer_elig_membs_full_name(memb_count), 		20, emer_row, fn_col
				EMReadScreen emer_elig_membs_request_yn(memb_count), 		1, emer_row, rq_col
				EMReadScreen emer_elig_membs_code(memb_count), 				1, emer_row, mc_col
				If emer_program = "EA" Then EMReadScreen emer_elig_membs_fund_fact(memb_count), 		1, emer_row, ff_col
				EMReadScreen emer_elig_membs_adult_or_child(memb_count), 	1, emer_row, ac_col
				EMReadScreen emer_elig_membs_elig_status(memb_count), 		10, emer_row, es_col

				If emer_elig_membs_code(memb_count) = "A" Then emer_elig_membs_info(memb_count) = "Counted Eligible"
				If emer_elig_membs_code(memb_count) = "F" Then emer_elig_membs_info(memb_count) = "Counted Ineligible"
				If emer_elig_membs_code(memb_count) = "N" Then emer_elig_membs_info(memb_count) = "Not Counted Ineligible"

				If emer_elig_membs_adult_or_child(memb_count) = "A" Then emer_elig_membs_adult_or_child(memb_count) = "Adult"
				If emer_elig_membs_adult_or_child(memb_count) = "C" Then emer_elig_membs_adult_or_child(memb_count) = "Child"

				emer_elig_membs_full_name(memb_count) = trim((emer_elig_membs_full_name(memb_count)))
				emer_elig_membs_elig_status(memb_count) = trim((emer_elig_membs_elig_status(memb_count)))

				If emer_program = "EA" Then EMWriteScreen "X", emer_row, 4

				memb_count = memb_count + 1
				emer_row = emer_row + 1
				EMReadScreen next_ref_numb, 2, emer_row, 6
			Loop until next_ref_numb = "  "

			transmit
			If emer_program = "EA" Then
				Do
					EMReadScreen person_name, 20, 18, 18
					person_name = trim(person_name)
					For each_memb = 0 to UBound(emer_elig_ref_numbs)
						If emer_elig_membs_full_name(each_memb) = person_name Then
							EMReadScreen emer_elig_membs_12_month_test(each_memb), 6, 13, 26
							EMReadScreen emer_elig_membs_last_emer_begin_date(each_memb), 8, 15, 29

							emer_elig_membs_12_month_test(each_memb) = trim(emer_elig_membs_12_month_test(each_memb))
							emer_elig_membs_last_emer_begin_date(each_memb) = trim(emer_elig_membs_last_emer_begin_date(each_memb))
						End If
					Next

					transmit
					EMReadScreen emer_panel, 4, 3, 49
				Loop until emer_panel = "EMCR"

				EMReadScreen emer_elig_case_test_citizenship, 		6, 8, 14
				EMReadScreen emer_elig_case_test_coop_MFIP, 		6, 9, 14
				EMReadScreen emer_elig_case_test_copayment, 		6, 10, 14
				EMReadScreen emer_elig_case_test_cost_effective, 	6, 11, 14
				EMReadScreen emer_elig_case_test_eligible_child, 	6, 12, 14
				EMReadScreen emer_elig_case_test_emergency, 		6, 13, 14

				EMReadScreen emer_elig_case_test_equitable_interest, 6, 8, 48
				EMReadScreen emer_elig_case_test_residency, 		6, 9, 48
				EMReadScreen emer_elig_case_test_resources, 		6, 10, 48
				EMReadScreen emer_elig_case_test_verif, 			6, 11, 48
				EMReadScreen emer_elig_case_test_12_month, 			6, 12, 48

				emer_elig_case_test_citizenship = trim(emer_elig_case_test_citizenship)
				emer_elig_case_test_coop_MFIP = trim(emer_elig_case_test_coop_MFIP)
				emer_elig_case_test_copayment = trim(emer_elig_case_test_copayment)
				emer_elig_case_test_cost_effective = trim(emer_elig_case_test_cost_effective)
				emer_elig_case_test_eligible_child = trim(emer_elig_case_test_eligible_child)
				emer_elig_case_test_emergency = trim(emer_elig_case_test_emergency)

				emer_elig_case_test_equitable_interest = trim(emer_elig_case_test_equitable_interest)
				emer_elig_case_test_residency = trim(emer_elig_case_test_residency)
				emer_elig_case_test_resources = trim(emer_elig_case_test_resources)
				emer_elig_case_test_verif = trim(emer_elig_case_test_verif)
				emer_elig_case_test_12_month = trim(emer_elig_case_test_12_month)
			End If

			If emer_program = "EG" Then
				EMReadScreen emer_elig_case_test_coop_work, 		6, 9, 7
				EMReadScreen emer_elig_case_test_copayment, 		6, 10, 7
				EMReadScreen emer_elig_case_test_cost_effective, 	6, 11, 7
				EMReadScreen emer_elig_case_test_county_allocation, 6, 12, 7
				EMReadScreen emer_elig_case_test_elig_other_program,6, 13, 7
				EMReadScreen emer_elig_case_test_emergency, 		6, 14, 7

				EMReadScreen emer_elig_case_test_equitable_interest, 6, 9, 49
				EMReadScreen emer_elig_case_test_resources, 		6, 10, 49
				EMReadScreen emer_elig_case_test_residency, 		6, 11, 49
				EMReadScreen emer_elig_case_test_verif, 			6, 12, 49
				EMReadScreen emer_elig_case_test_12_month, 			6, 13, 49
				EMReadScreen emer_elig_case_test_200_percent_fpg, 	6, 14, 49

				emer_elig_case_test_coop_work = trim(emer_elig_case_test_coop_work)
				emer_elig_case_test_copayment = trim(emer_elig_case_test_copayment)
				emer_elig_case_test_cost_effective = trim(emer_elig_case_test_cost_effective)
				emer_elig_case_test_county_allocation = trim(emer_elig_case_test_county_allocation)
				emer_elig_case_test_elig_other_program = trim(emer_elig_case_test_elig_other_program)
				emer_elig_case_test_emergency = trim(emer_elig_case_test_emergency)

				emer_elig_case_test_equitable_interest = trim(emer_elig_case_test_equitable_interest)
				emer_elig_case_test_resources = trim(emer_elig_case_test_resources)
				emer_elig_case_test_residency = trim(emer_elig_case_test_residency)
				emer_elig_case_test_verif = trim(emer_elig_case_test_verif)
				emer_elig_case_test_12_month = trim(emer_elig_case_test_12_month)
				emer_elig_case_test_200_percent_fpg = trim(emer_elig_case_test_200_percent_fpg)
			End If

			transmit 		'going to EMAV'

			EMReadScreen emer_elig_available_gross_earned_income, 	9, 7, 32
			EMReadScreen emer_elig_available_actual_work_expense, 	9, 8, 32
			EMReadScreen emer_elig_available_net_earned_income, 	9, 9, 32

			EMReadScreen emer_elig_available_unearned_income, 		9, 7, 71
			EMReadScreen emer_elig_available_assets, 				9, 8, 17
			EMReadScreen emer_elig_available_other_assets, 			9, 9, 71
			EMReadScreen emer_elig_available_total_income_assets, 	9, 10, 71

			EMReadScreen emer_elig_expense_rent_mortgage, 		9, 13, 32
			EMReadScreen emer_elig_expense_fuel,	 			9, 14, 32
			EMReadScreen emer_elig_expense_electric, 			9, 15, 32
			EMReadScreen emer_elig_expense_msa_standard, 		9, 16, 32

			EMReadScreen emer_elig_expense_car_payment, 		9, 13, 71
			EMReadScreen emer_elig_expense_phone, 				9, 14, 71
			EMReadScreen emer_elig_expense_food, 				9, 15, 71
			EMReadScreen emer_elig_expense_other, 				9, 16, 71
			EMReadScreen emer_elig_total_basic_needs, 			9, 17, 71
			EMReadScreen emer_elig_expense_net_income_assets, 	9, 18, 71

			emer_elig_available_gross_earned_income = trim(emer_elig_available_gross_earned_income)
			emer_elig_available_actual_work_expense = trim(emer_elig_available_actual_work_expense)
			emer_elig_available_net_earned_income = trim(emer_elig_available_net_earned_income)
			emer_elig_available_unearned_income = trim(emer_elig_available_unearned_income)
			emer_elig_available_assets = trim(emer_elig_available_assets)
			emer_elig_available_other_assets = trim(emer_elig_available_other_assets)
			emer_elig_available_total_income_assets = trim(emer_elig_available_total_income_assets)
			emer_elig_expense_rent_mortgage = trim(emer_elig_expense_rent_mortgage)
			emer_elig_expense_fuel = trim(emer_elig_expense_fuel)
			emer_elig_expense_electric = trim(emer_elig_expense_electric)
			emer_elig_expense_msa_standard = trim(emer_elig_expense_msa_standard)
			emer_elig_expense_car_payment = trim(emer_elig_expense_car_payment)
			emer_elig_expense_phone = trim(emer_elig_expense_phone)
			emer_elig_expense_food = trim(emer_elig_expense_food)
			emer_elig_expense_other = trim(emer_elig_expense_other)
			emer_elig_total_basic_needs = trim(emer_elig_total_basic_needs)
			emer_elig_expense_net_income_assets = trim(emer_elig_expense_net_income_assets)

			emer_elig_available_gross_earned_income = replace(emer_elig_available_gross_earned_income, "_", "")
			emer_elig_available_actual_work_expense = replace(emer_elig_available_actual_work_expense, "_", "")
			emer_elig_available_net_earned_income = replace(emer_elig_available_net_earned_income, "_", "")
			emer_elig_available_unearned_income = replace(emer_elig_available_unearned_income, "_", "")
			emer_elig_available_assets = replace(emer_elig_available_assets, "_", "")
			emer_elig_available_other_assets = replace(emer_elig_available_other_assets, "_", "")
			emer_elig_available_total_income_assets = replace(emer_elig_available_total_income_assets, "_", "")
			emer_elig_expense_rent_mortgage = replace(emer_elig_expense_rent_mortgage, "_", "")
			emer_elig_expense_fuel = replace(emer_elig_expense_fuel, "_", "")
			emer_elig_expense_electric = replace(emer_elig_expense_electric, "_", "")
			emer_elig_expense_msa_standard = replace(emer_elig_expense_msa_standard, "_", "")
			emer_elig_expense_car_payment = replace(emer_elig_expense_car_payment, "_", "")
			emer_elig_expense_phone = replace(emer_elig_expense_phone, "_", "")
			emer_elig_expense_food = replace(emer_elig_expense_food, "_", "")
			emer_elig_expense_other = replace(emer_elig_expense_other, "_", "")
			emer_elig_total_basic_needs = replace(emer_elig_total_basic_needs, "_", "")
			emer_elig_expense_net_income_assets = replace(emer_elig_expense_net_income_assets, "_", "")

			transmit 'go to EMSM'

			EMReadScreen emer_elig_approved_date, 			8, 3, 14
			EMReadScreen emer_elig_process_date, 			8, 2, 73
			EMReadScreen emer_elig_summ_date_last_approval, 8, 6, 32
			EMReadScreen emer_elig_summ_current_program_status, 10, 7, 32
			EMReadScreen emer_elig_summ_eligibility_result, 10, 8, 32
			EMReadScreen emer_elig_summ_last_used, 			8, 9, 32

			EMReadScreen emer_elig_summ_adults_in_unit, 	2, 6, 73
			EMReadScreen emer_elig_summ_children_in_unit, 	2, 7, 73
			EMReadScreen emer_elig_summ_begin_date, 		8, 8, 71
			EMReadScreen emer_elig_summ_end_date, 			8, 9, 71

			EMReadScreen emer_elig_summ_need_foreclosure, 	9, 11, 32
			EMReadScreen emer_elig_summ_need_temp_shelter, 	9, 12, 32
			EMReadScreen emer_elig_summ_need_other_shelter, 9, 13, 32
			EMReadScreen emer_elig_summ_need_utility, 		9, 14, 32
			EMReadScreen emer_elig_summ_need_other, 		9, 15, 32
			EMReadScreen emer_elig_summ_need_total, 		9, 16, 32

			EMReadScreen emer_elig_summ_payment, 			9, 13, 71

			emer_elig_summ_date_last_approval = trim(emer_elig_summ_date_last_approval)
			emer_elig_summ_current_program_status = trim(emer_elig_summ_current_program_status)
			emer_elig_summ_last_used = trim(emer_elig_summ_last_used)

			emer_elig_summ_adults_in_unit = trim(emer_elig_summ_adults_in_unit)
			emer_elig_summ_children_in_unit = trim(emer_elig_summ_children_in_unit)

			emer_elig_summ_eligibility_result = replace(emer_elig_summ_eligibility_result, "_", "")

			emer_elig_summ_need_foreclosure = replace(emer_elig_summ_need_foreclosure, "_", "")
			emer_elig_summ_need_temp_shelter = replace(emer_elig_summ_need_temp_shelter, "_", "")
			emer_elig_summ_need_other_shelter = replace(emer_elig_summ_need_other_shelter, "_", "")
			emer_elig_summ_need_utility = replace(emer_elig_summ_need_utility, "_", "")
			emer_elig_summ_need_other = replace(emer_elig_summ_need_other, "_", "")

			emer_elig_summ_need_total = trim(emer_elig_summ_need_total)
			emer_elig_summ_payment = trim(emer_elig_summ_payment)

			''TODO - open foreclosure and utility pop-up
		End If

		Call back_to_SELF
	end sub


end class


class snap_eligibility_detail

	public elig_footer_month
	public elig_footer_year
	public elig_version_number
	public elig_version_date
	public elig_version_result
	public approved_today
	public approved_version_found
	public er_month
	public hrf_month
	public er_status
	public er_caf_date
	public er_interview_date
	public hrf_status
	public hrf_doc_date

	public snap_elig_ref_numbs()
	public snap_elig_membs_request_yn()
	public snap_elig_membs_code()
	public snap_elig_membs_status_info()
	public snap_elig_membs_counted()
	public snap_elig_membs_state_food()
	public snap_elig_membs_eligibility()
	public snap_elig_membs_begin_date()
	public snap_elig_membs_budget_cycle()

	public snap_elig_membs_abawd()
	public snap_elig_membs_absence()
	public snap_elig_membs_roomer()
	public snap_elig_membs_boarder()
	public snap_elig_membs_citizenship()
	public snap_elig_membs_citizenship_coop()
	public snap_elig_membs_cmdty()
	public snap_elig_membs_disq()
	public snap_elig_membs_dupl_assist()
	public snap_elig_membs_fraud()
	public snap_elig_membs_eligible_student()
	public snap_elig_membs_institution()
	public snap_elig_membs_mfip_elig()
	public snap_elig_membs_non_applcnt()
	public snap_elig_membs_residence()
	public snap_elig_membs_ssn_coop()
	public snap_elig_membs_unit_memb()
	public snap_elig_membs_work_reg()
	public snap_elig_membs_failed_test()
	public snap_elig_membs_drug_felon_test()

	public snap_expedited
	public snap_uhfs
	public snap_exp_package_includes_month_one
	public snap_exp_package_includes_month_two
	public elig_membs_list
	public inelig_membs_list
	public snap_prorated
	public snap_earned_income_budgeted
	public snap_unearned_income_budgeted
	public snap_shel_costs_budgeted
	public snap_hest_costs_budgeted
	public snap_categorical_eligibility
	public snap_case_appl_withdrawn_test
	public snap_case_applct_elig_test
	public snap_case_comdty_test
	public snap_case_disq_test
	public snap_case_dupl_assist_test
	public snap_case_eligible_person_test
	public snap_case_fail_coop_test
	public snap_case_fail_file_test
	public snap_case_prosp_gross_inc_test
	public snap_case_prosp_net_inc_test
	public snap_case_recert_test
	public snap_case_residence_test
	public snap_case_resource_test
	public snap_case_retro_gross_inc_test
	public snap_case_retro_net_inc_test
	public snap_case_strike_test
	public snap_case_xfer_resource_inc_test
	public snap_case_verif_test
	public snap_case_verif_test_MEMB_ID
	public snap_case_verif_test_ACCT
	public snap_case_verif_test_PACT
	public snap_case_verif_test_ADDR
	public snap_case_verif_test_SECU
	public snap_case_verif_test_RBIC
	public snap_case_verif_test_BUSI
	public snap_case_verif_test_SPON
	public snap_case_verif_test_STIN
	public snap_case_verif_test_UNEA
	public snap_case_verif_test_JOBS
	public snap_case_verif_test_STWK
	public snap_case_verif_test_STRK
	public snap_case_voltry_quit_test
	public snap_case_work_reg_test
	public snap_fail_file_hrf
	public snap_fail_file_sr
	public snap_resource_cash
	public snap_resource_acct
	public snap_resource_secu
	public snap_resource_cars
	public snap_resource_rest
	public snap_resource_other
	public snap_resource_burial
	public snap_resource_spon
	public snap_resource_total
	public snap_resource_max
	public snap_budg_gross_wages
	public snap_budg_self_emp
	public snap_budg_total_earned_inc
	public snap_budg_pa_grant_inc
	public snap_budg_rsdi_inc
	public snap_budg_ssi_inc
	public snap_budg_va_inc
	public snap_budg_uc_wc_inc
	public snap_budg_cses_inc
	public snap_budg_other_unea_inc
	public snap_budg_total_unea_inc
	public snap_budg_schl_inc
	public snap_budg_farm_ofset
	public snap_budg_total_gross_inc
	public snap_budg_max_gross_inc
	public snap_budg_deduct_standard
	public snap_budg_deduct_earned
	public snap_budg_deduct_medical
	public snap_budg_deduct_depndt_care
	public snap_budg_deduct_cses
	public snap_budg_total_deduct
	public snap_budg_net_inc
	public snap_budg_shel_rent_mort
	public snap_budg_shel_prop_tax
	public snap_budg_shel_home_ins
	public snap_budg_shel_electricity
	public snap_budg_shel_heat_ac
	public snap_budg_shel_water_garbage
	public snap_budg_shel_phone
	public snap_budg_shel_other
	public snap_budg_housing_exp_total
	public snap_budg_utilities_exp_total
	public snap_budg_utilities_list
	public snap_budg_shel_total
	public snap_budg_50_perc_net_inc
	public snap_budg_adj_shel_costs
	public snap_budg_max_allow_shel
	public snap_budg_shel_expenses
	' public snap_budg_net_adj_inc
	public snap_budg_max_net_adj_inc
	public snap_benefit_monthly_fs_allot
	public snap_benefit_drug_felon_sanc_amt
	public snap_benefit_amt_already_issued
	public snap_benefit_recoup_amount
	public snap_benefit_benefit_amount
	public snap_benefit_state_food_amt
	public snap_benefit_fed_food_amt
	public snap_benefit_recoup_from_fed_fs
	public snap_benefit_recoup_from_state_fs
	public snap_approved_date
	public snap_date_last_approval
	public snap_curr_prog_status
	public snap_elig_result
	public snap_reporting_status
	public snap_info_source
	public snap_benefit
	public snap_elig_revw_date
	public snap_budget_cycle
	public snap_budg_numb_in_assist_unit
	public adults_recv_snap
	public children_recv_snap
	public snap_budg_total_resources
	public snap_budg_max_resources
	public snap_budg_net_adj_inc
	public snap_bug_30_percent_net_adj_inc
	public snap_budg_thrifty_food_plan
	public snap_benefit_monthly_fs_allotment
	public snap_benefit_prorated_amt
	public snap_benefit_prorated_date
	public snap_benefit_amt
	public snap_exp_criteria_migrant_destitute
	public snap_exp_criteria_resource_100_income_150
	public snap_exp_criteria_resource_income_less_shelter
	public snap_exp_verif_status_postponed
	public snap_exp_verif_status_out_of_state
	public snap_exp_verif_status_all_provided
	public snap_elig_worker_message_one
	public snap_elig_worker_message_two


	public sub read_elig()
		approved_today = False
		approved_version_found = False

		snap_expedited = False
		snap_uhfs = False
		snap_exp_package_includes_month_one = False
		snap_exp_package_includes_month_two = False
		snap_prorated = False
		snap_earned_income_budgeted = False
		snap_unearned_income_budgeted = False
		snap_shel_costs_budgeted = False
		snap_hest_costs_budgeted = False
		snap_categorical_eligibility = ""

		ReDim snap_elig_ref_numbs(0)
		ReDim snap_elig_membs_request_yn(0)
		ReDim snap_elig_membs_code(0)
		ReDim snap_elig_membs_status_info(0)
		ReDim snap_elig_membs_counted(0)
		ReDim snap_elig_membs_state_food(0)
		ReDim snap_elig_membs_eligibility(0)
		ReDim snap_elig_membs_begin_date(0)
		ReDim snap_elig_membs_budget_cycle(0)
		ReDim snap_elig_membs_abawd(0)
		ReDim snap_elig_membs_absence(0)
		ReDim snap_elig_membs_roomer(0)
		ReDim snap_elig_membs_boarder(0)
		ReDim snap_elig_membs_citizenship(0)
		ReDim snap_elig_membs_citizenship_coop(0)
		ReDim snap_elig_membs_cmdty(0)
		ReDim snap_elig_membs_disq(0)
		ReDim snap_elig_membs_dupl_assist(0)
		ReDim snap_elig_membs_fraud(0)
		ReDim snap_elig_membs_eligible_student(0)
		ReDim snap_elig_membs_institution(0)
		ReDim snap_elig_membs_mfip_elig(0)
		ReDim snap_elig_membs_non_applcnt(0)
		ReDim snap_elig_membs_residence(0)
		ReDim snap_elig_membs_ssn_coop(0)
		ReDim snap_elig_membs_unit_memb(0)
		ReDim snap_elig_membs_work_reg(0)
		ReDim snap_elig_membs_failed_test(0)
		ReDim snap_elig_membs_drug_felon_test(0)

		call navigate_to_MAXIS_screen("ELIG", "FS  ")
		EMWriteScreen elig_footer_month, 19, 54
		EMWriteScreen elig_footer_year, 19, 57
		Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		' EMReadScreen approved_today, 8, 3, 14
		' approved_today = DateAdd("d", 0, approved_today)
		If approved_version_found = True Then
			If DateDiff("d", date, elig_version_date) = 0 Then approved_today = True

			row = 7
			elig_memb_count = 0
			elig_membs_list = ""
			inelig_membs_list = ""
			Do
				EMReadScreen ref_numb, 2, row, 10

				ReDim preserve snap_elig_ref_numbs(elig_memb_count)
				ReDim preserve snap_elig_membs_request_yn(elig_memb_count)
				ReDim preserve snap_elig_membs_code(elig_memb_count)
				ReDim preserve snap_elig_membs_status_info(elig_memb_count)
				ReDim preserve snap_elig_membs_counted(elig_memb_count)
				ReDim preserve snap_elig_membs_state_food(elig_memb_count)
				ReDim preserve snap_elig_membs_eligibility(elig_memb_count)
				ReDim preserve snap_elig_membs_begin_date(elig_memb_count)
				ReDim preserve snap_elig_membs_budget_cycle(elig_memb_count)

				ReDim preserve snap_elig_membs_abawd(elig_memb_count)
				ReDim preserve snap_elig_membs_absence(elig_memb_count)
				ReDim preserve snap_elig_membs_roomer(elig_memb_count)
				ReDim preserve snap_elig_membs_boarder(elig_memb_count)
				ReDim preserve snap_elig_membs_citizenship(elig_memb_count)
				ReDim preserve snap_elig_membs_citizenship_coop(elig_memb_count)
				ReDim preserve snap_elig_membs_cmdty(elig_memb_count)
				ReDim preserve snap_elig_membs_disq(elig_memb_count)
				ReDim preserve snap_elig_membs_dupl_assist(elig_memb_count)
				ReDim preserve snap_elig_membs_fraud(elig_memb_count)
				ReDim preserve snap_elig_membs_eligible_student(elig_memb_count)
				ReDim preserve snap_elig_membs_institution(elig_memb_count)
				ReDim preserve snap_elig_membs_mfip_elig(elig_memb_count)
				ReDim preserve snap_elig_membs_non_applcnt(elig_memb_count)
				ReDim preserve snap_elig_membs_residence(elig_memb_count)
				ReDim preserve snap_elig_membs_ssn_coop(elig_memb_count)
				ReDim preserve snap_elig_membs_unit_memb(elig_memb_count)
				ReDim preserve snap_elig_membs_work_reg(elig_memb_count)
				ReDim preserve snap_elig_membs_failed_test(elig_memb_count)
				ReDim preserve snap_elig_membs_drug_felon_test(elig_memb_count)

				snap_elig_ref_numbs(elig_memb_count) = ref_numb
				EMReadScreen snap_elig_membs_request_yn(elig_memb_count), 1, row, 32
				EMReadScreen snap_elig_membs_code(elig_memb_count), 1, row, 35
				EMReadScreen memb_count, 11, row, 39
				EMReadScreen memb_state_food, 1, row, 50
				EMReadScreen memb_elig, 10, row, 57
				EMReadScreen snap_elig_membs_begin_date(elig_memb_count), 8, row, 68
				EMReadScreen memb_budg_cycle, 1, row, 78

				If snap_elig_membs_code(elig_memb_count) = "A" Then snap_elig_membs_status_info(elig_memb_count) = "Eligible"
				If snap_elig_membs_code(elig_memb_count) = "C" Then snap_elig_membs_status_info(elig_memb_count) = "Citizenship"
				If snap_elig_membs_code(elig_memb_count) = "F" Then snap_elig_membs_status_info(elig_memb_count) = "Fraud, DISQ, Work Reg"
				If snap_elig_membs_code(elig_memb_count) = "D" Then snap_elig_membs_status_info(elig_memb_count) = "Duplicate Assistance"
				If snap_elig_membs_code(elig_memb_count) = "I" Then snap_elig_membs_status_info(elig_memb_count) = "Ineligible"
				If snap_elig_membs_code(elig_memb_count) = "N" Then snap_elig_membs_status_info(elig_memb_count) = "Unit Member"
				If snap_elig_membs_code(elig_memb_count) = "S" Then snap_elig_membs_status_info(elig_memb_count) = "Ineligible Student"
				If snap_elig_membs_code(elig_memb_count) = "U" Then snap_elig_membs_status_info(elig_memb_count) = "Unknown"
				snap_elig_membs_counted(elig_memb_count) = trim(memb_count)
				If memb_state_food = "Y" Then snap_elig_membs_state_food(elig_memb_count) = True
				If memb_state_food = "N" Then snap_elig_membs_state_food(elig_memb_count) = False
				snap_elig_membs_eligibility(elig_memb_count) = trim(memb_elig)
				If memb_budg_cycle = "P" Then snap_elig_membs_budget_cycle(elig_memb_count) = "Prospective"
				If memb_budg_cycle = "R" Then snap_elig_membs_budget_cycle(elig_memb_count) = "Retrospective"
				If snap_elig_membs_eligibility(elig_memb_count) = "ELIGIBLE" Then elig_membs_list = elig_membs_list & "Memb " & snap_elig_ref_numbs(elig_memb_count) & ", "
				If snap_elig_membs_eligibility(elig_memb_count) = "INELIGIBLE" Then inelig_membs_list = inelig_membs_list & "Memb " & snap_elig_ref_numbs(elig_memb_count) & ", "

				Call write_value_and_transmit("X", row, 5)

				EMReadScreen memb_abawd, 			6, 6, 20
				EMReadScreen memb_absence, 			6, 7, 20
				EMReadScreen memb_roomer, 			6, 8, 20
				EMReadScreen memb_boarder, 			6, 9, 20
				EMReadScreen memb_citizenship, 		6, 10, 20
				EMReadScreen memb_citizenship_coop, 6, 11, 20
				EMReadScreen memb_cmdty, 			6, 12, 20
				EMReadScreen memb_disq,				6, 13, 20
				EMReadScreen memb_dupl_assist, 		6, 14, 20

				snap_elig_membs_abawd(elig_memb_count) = trim(memb_abawd)
				snap_elig_membs_absence(elig_memb_count) = trim(memb_absence)
				snap_elig_membs_roomer(elig_memb_count) = trim(memb_roomer)
				snap_elig_membs_boarder(elig_memb_count) = trim(memb_boarder)
				snap_elig_membs_citizenship(elig_memb_count) = trim(memb_citizenship)
				snap_elig_membs_citizenship_coop(elig_memb_count) = trim(memb_citizenship_coop)
				snap_elig_membs_cmdty(elig_memb_count) = trim(memb_cmdty)
				snap_elig_membs_disq(elig_memb_count) = trim(memb_disq)
				snap_elig_membs_dupl_assist(elig_memb_count) = trim(memb_dupl_assist)

				EMReadScreen memb_fraud, 			6, 6, 54
				EMReadScreen memb_eligible_student, 6, 7, 54
				EMReadScreen memb_institution, 		6, 8, 54
				EMReadScreen memb_mfip_elig, 		6, 9, 54
				EMReadScreen memb_non_applcnt, 		6, 10, 54
				EMReadScreen memb_residence, 		6, 11, 54
				EMReadScreen memb_ssn_coop, 		6, 12, 54
				EMReadScreen memb_unit_memb, 		6, 13, 54
				EMReadScreen memb_work_reg, 		6, 14, 54

				snap_elig_membs_fraud(elig_memb_count) = trim(memb_fraud)
				snap_elig_membs_eligible_student(elig_memb_count) = trim(memb_eligible_student)
				snap_elig_membs_institution(elig_memb_count) = trim(memb_institution)
				snap_elig_membs_mfip_elig(elig_memb_count) = trim(memb_mfip_elig)
				snap_elig_membs_non_applcnt(elig_memb_count) = trim(memb_non_applcnt)
				snap_elig_membs_residence(elig_memb_count) = trim(memb_residence)
				snap_elig_membs_ssn_coop(elig_memb_count) = trim(memb_ssn_coop)
				snap_elig_membs_unit_memb(elig_memb_count) = trim(memb_unit_memb)
				snap_elig_membs_work_reg(elig_memb_count) = trim(memb_work_reg)

				snap_elig_membs_failed_test(elig_memb_count) = False

				If snap_elig_membs_abawd(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_absence(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_roomer(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_boarder(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_citizenship(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_citizenship_coop(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_cmdty(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_disq(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_dupl_assist(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_fraud(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_eligible_student(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_institution(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_mfip_elig(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_non_applcnt(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_residence(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_ssn_coop(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_unit_memb(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				If snap_elig_membs_work_reg(elig_memb_count) = "FAILED" Then snap_elig_membs_failed_test(elig_memb_count) = True
				transmit


				elig_memb_count = elig_memb_count + 1
				row = row + 1
				EMReadScreen next_ref_numb, 2, row, 10
			Loop until next_ref_numb = "  "

			elig_membs_list = trim(elig_membs_list)
			inelig_membs_list = trim(inelig_membs_list)
			If right(elig_membs_list, 1) = "," Then elig_membs_list = left(elig_membs_list, len(elig_membs_list)-1)
			If right(inelig_membs_list, 1) = "," Then inelig_membs_list = left(inelig_membs_list, len(inelig_membs_list)-1)

			transmit 		'FSCR
			EMReadScreen case_expedited_indicator, 9, 4, 3
			If case_expedited_indicator = "EXPEDITED" Then snap_expedited = True
			EMReadScreen case_uhfs_indicator, 11, 5, 4
			If case_uhfs_indicator = "UNCLE HARRY" Then snap_uhfs = True

			EMReadScreen snap_case_appl_withdrawn_test, 	6, 7, 9
			EMReadScreen snap_case_applct_elig_test, 		6, 8, 9
			EMReadScreen snap_case_comdty_test, 			6, 9, 9
			EMReadScreen snap_case_disq_test, 				6, 10, 9
			EMReadScreen snap_case_dupl_assist_test, 		6, 11, 9
			EMReadScreen snap_case_eligible_person_test, 	6, 12, 9
			EMReadScreen snap_case_fail_coop_test, 			6, 13, 9
			EMReadScreen snap_case_fail_file_test, 			6, 14, 9
			EMReadScreen snap_case_prosp_gross_inc_test, 	6, 15, 9
			EMReadScreen snap_case_prosp_net_inc_test, 		6, 16, 9
			snap_case_appl_withdrawn_test = trim(snap_case_appl_withdrawn_test)
			snap_case_applct_elig_test = trim(snap_case_applct_elig_test)
			snap_case_comdty_test = trim(snap_case_comdty_test)
			snap_case_disq_test = trim(snap_case_disq_test)
			snap_case_dupl_assist_test = trim(snap_case_dupl_assist_test)
			snap_case_eligible_person_test = trim(snap_case_eligible_person_test)
			snap_case_fail_coop_test = trim(snap_case_fail_coop_test)
			snap_case_fail_file_test = trim(snap_case_fail_file_test)
			snap_case_prosp_gross_inc_test = trim(snap_case_prosp_gross_inc_test)
			snap_case_prosp_net_inc_test = trim(snap_case_prosp_net_inc_test)

			EMReadScreen snap_case_recert_test, 			6, 7, 49
			EMReadScreen snap_case_residence_test, 			6, 8, 49
			EMReadScreen snap_case_resource_test, 			6, 9, 49
			EMReadScreen snap_case_retro_gross_inc_test, 	6, 10, 49
			EMReadScreen snap_case_retro_net_inc_test, 		6, 11, 49
			EMReadScreen snap_case_strike_test, 			6, 12, 49
			EMReadScreen snap_case_xfer_resource_inc_test, 	6, 13, 49
			EMReadScreen snap_case_verif_test, 				6, 14, 49
			EMReadScreen snap_case_voltry_quit_test, 		6, 15, 49
			EMReadScreen snap_case_work_reg_test, 			6, 16, 49
			snap_case_recert_test = trim(snap_case_recert_test)
			snap_case_residence_test = trim(snap_case_residence_test)
			snap_case_resource_test = trim(snap_case_resource_test)
			snap_case_retro_gross_inc_test = trim(snap_case_retro_gross_inc_test)
			snap_case_retro_net_inc_test = trim(snap_case_retro_net_inc_test)
			snap_case_strike_test = trim(snap_case_strike_test)
			snap_case_xfer_resource_inc_test = trim(snap_case_xfer_resource_inc_test)
			snap_case_verif_test = trim(snap_case_verif_test)
			snap_case_voltry_quit_test = trim(snap_case_voltry_quit_test)
			snap_case_work_reg_test = trim(snap_case_work_reg_test)

			Call write_value_and_transmit("X", 14, 4)		''Fail to File Detail
			EMReadScreen snap_fail_file_hrf, 6, 10, 32
			EMReadScreen snap_fail_file_sr, 6, 11, 32
			transmit
			snap_fail_file_hrf = trim(snap_fail_file_hrf)
			snap_fail_file_sr = trim(snap_fail_file_sr)

			Call write_value_and_transmit("X", 14, 4)		''Resource Detail
			EMReadScreen snap_resource_cash, 	10, 8, 47
			EMReadScreen snap_resource_acct, 	10, 9, 47
			EMReadScreen snap_resource_secu, 	10, 10, 47
			EMReadScreen snap_resource_cars, 	10, 11, 47
			EMReadScreen snap_resource_rest, 	10, 12, 47
			EMReadScreen snap_resource_other, 	10, 13, 47
			EMReadScreen snap_resource_burial, 	10, 14, 47
			EMReadScreen snap_resource_spon, 	10, 15, 47
			EMReadScreen snap_resource_total, 	10, 17, 47
			EMReadScreen snap_resource_max, 	10, 18, 47
			transmit

			snap_resource_cash = trim(snap_resource_cash)
			snap_resource_acct = trim(snap_resource_acct)
			snap_resource_secu = trim(snap_resource_secu)
			snap_resource_cars = trim(snap_resource_cars)
			snap_resource_rest = trim(snap_resource_rest)
			snap_resource_other = trim(snap_resource_other)
			snap_resource_burial = trim(snap_resource_burial)
			snap_resource_spon = trim(snap_resource_spon)
			snap_resource_total = trim(snap_resource_total)
			snap_resource_max = trim(snap_resource_max)

			If snap_case_verif_test = "FAILED" Then
				Call write_value_and_transmit("X", 14, 44)
				EMReadScreen snap_case_verif_test_MEMB_ID, 6, 7, 30
				EMReadScreen snap_case_verif_test_ACCT, 6, 8, 30
				EMReadScreen snap_case_verif_test_PACT, 6, 9, 30
				EMReadScreen snap_case_verif_test_ADDR, 6, 10, 30
				EMReadScreen snap_case_verif_test_SECU, 6, 11, 30
				EMReadScreen snap_case_verif_test_RBIC, 6, 12, 30
				EMReadScreen snap_case_verif_test_BUSI, 6, 13, 30
				EMReadScreen snap_case_verif_test_SPON, 6, 14, 30
				EMReadScreen snap_case_verif_test_STIN, 6, 15, 30
				EMReadScreen snap_case_verif_test_UNEA, 6, 16, 30
				EMReadScreen snap_case_verif_test_JOBS, 6, 17, 30
				EMReadScreen snap_case_verif_test_STWK, 6, 18, 30
				EMReadScreen snap_case_verif_test_STRK, 6, 19, 30
				transmit

				snap_case_verif_test_MEMB_ID = trim(snap_case_verif_test_MEMB_ID)
				snap_case_verif_test_ACCT = trim(snap_case_verif_test_ACCT)
				snap_case_verif_test_PACT = trim(snap_case_verif_test_PACT)
				snap_case_verif_test_ADDR = trim(snap_case_verif_test_ADDR)
				snap_case_verif_test_SECU = trim(snap_case_verif_test_SECU)
				snap_case_verif_test_RBIC = trim(snap_case_verif_test_RBIC)
				snap_case_verif_test_BUSI = trim(snap_case_verif_test_BUSI)
				snap_case_verif_test_SPON = trim(snap_case_verif_test_SPON)
				snap_case_verif_test_STIN = trim(snap_case_verif_test_STIN)
				snap_case_verif_test_UNEA = trim(snap_case_verif_test_UNEA)
				snap_case_verif_test_JOBS = trim(snap_case_verif_test_JOBS)
				snap_case_verif_test_STWK = trim(snap_case_verif_test_STWK)
				snap_case_verif_test_STRK = trim(snap_case_verif_test_STRK)
			End If

			transmit 		'FSB1
			EMReadScreen snap_budg_gross_wages, 		10, 5, 31
			EMReadScreen snap_budg_self_emp, 			10, 6, 31
			EMReadScreen snap_budg_total_earned_inc, 	10, 8, 31

			snap_budg_gross_wages = trim(snap_budg_gross_wages)
			snap_budg_self_emp = trim(snap_budg_self_emp)
			snap_budg_total_earned_inc = trim(snap_budg_total_earned_inc)
			If snap_budg_total_earned_inc = "" Then snap_budg_total_earned_inc = "0.00"


			EMReadScreen snap_budg_pa_grant_inc, 	10, 10, 31
			EMReadScreen snap_budg_rsdi_inc, 		10, 11, 31
			EMReadScreen snap_budg_ssi_inc, 		10, 12, 31
			EMReadScreen snap_budg_va_inc, 			10, 13, 31
			EMReadScreen snap_budg_uc_wc_inc, 		10, 14, 31
			EMReadScreen snap_budg_cses_inc, 		10, 15, 31
			EMReadScreen snap_budg_other_unea_inc, 	10, 16, 31
			EMReadScreen snap_budg_total_unea_inc, 	10, 18, 31

			snap_budg_pa_grant_inc = trim(snap_budg_pa_grant_inc)
			snap_budg_rsdi_inc = trim(snap_budg_rsdi_inc)
			snap_budg_ssi_inc = trim(snap_budg_ssi_inc)
			snap_budg_va_inc = trim(snap_budg_va_inc)
			snap_budg_uc_wc_inc = trim(snap_budg_uc_wc_inc)
			snap_budg_cses_inc = trim(snap_budg_cses_inc)
			snap_budg_other_unea_inc = trim(snap_budg_other_unea_inc)
			snap_budg_total_unea_inc = trim(snap_budg_total_unea_inc)
			If snap_budg_total_unea_inc = "" Then snap_budg_total_unea_inc = "0.00"

			EMReadScreen snap_budg_schl_inc, 			10, 5, 71
			EMReadScreen snap_budg_farm_ofset, 			10, 6, 71
			EMReadScreen snap_budg_total_gross_inc, 	10, 7, 71
			EMReadScreen snap_budg_max_gross_inc, 		10, 8, 71

			EMReadScreen snap_budg_deduct_standard, 	10, 10, 71
			EMReadScreen snap_budg_deduct_earned, 		10, 11, 71
			EMReadScreen snap_budg_deduct_medical, 		10, 12, 71
			EMReadScreen snap_budg_deduct_depndt_care, 	10, 13, 71
			EMReadScreen snap_budg_deduct_cses, 		10, 14, 71
			EMReadScreen snap_budg_total_deduct, 		10, 16, 71

			EMReadScreen snap_budg_net_inc, 			10, 18, 71

			snap_budg_schl_inc = trim(snap_budg_schl_inc)
			snap_budg_farm_ofset = trim(snap_budg_farm_ofset)
			snap_budg_total_gross_inc = trim(snap_budg_total_gross_inc)
			snap_budg_max_gross_inc = trim(snap_budg_max_gross_inc)
			snap_budg_deduct_standard = trim(snap_budg_deduct_standard)
			snap_budg_deduct_earned = trim(snap_budg_deduct_earned)
			snap_budg_deduct_medical = trim(snap_budg_deduct_medical)
			snap_budg_deduct_depndt_care = trim(snap_budg_deduct_depndt_care)
			snap_budg_deduct_cses = trim(snap_budg_deduct_cses)
			snap_budg_total_deduct = trim(snap_budg_total_deduct)
			snap_budg_net_inc = trim(snap_budg_net_inc)

			transmit 		'FSB2
			EMReadScreen snap_budg_shel_rent_mort, 		10, 5, 27
			EMReadScreen snap_budg_shel_prop_tax, 		10, 6, 27
			EMReadScreen snap_budg_shel_home_ins, 		10, 7, 27
			EMReadScreen snap_budg_shel_electricity, 	10, 8, 27
			EMReadScreen snap_budg_shel_heat_ac, 		10, 9, 27
			EMReadScreen snap_budg_shel_water_garbage, 	10, 10, 27
			EMReadScreen snap_budg_shel_phone, 			10, 11, 27
			EMReadScreen snap_budg_shel_other, 			10, 12, 27
			EMReadScreen snap_budg_shel_total, 			10, 14, 27
			EMReadScreen snap_budg_50_perc_net_inc, 	10, 15, 27
			EMReadScreen snap_budg_adj_shel_costs, 		10, 17, 27

			snap_budg_shel_rent_mort = trim(snap_budg_shel_rent_mort)
			snap_budg_shel_prop_tax = trim(snap_budg_shel_prop_tax)
			snap_budg_shel_home_ins = trim(snap_budg_shel_home_ins)
			snap_budg_shel_electricity = trim(snap_budg_shel_electricity)
			snap_budg_shel_heat_ac = trim(snap_budg_shel_heat_ac)
			snap_budg_shel_water_garbage = trim(snap_budg_shel_water_garbage)
			snap_budg_shel_phone = trim(snap_budg_shel_phone)
			snap_budg_shel_other = trim(snap_budg_shel_other)
			snap_budg_shel_total = trim(snap_budg_shel_total)
			snap_budg_50_perc_net_inc = trim(snap_budg_50_perc_net_inc)
			snap_budg_adj_shel_costs = trim(snap_budg_adj_shel_costs)

			If snap_budg_shel_rent_mort = "" Then snap_budg_shel_rent_mort = 0
			If snap_budg_shel_prop_tax = "" Then snap_budg_shel_prop_tax = 0
			If snap_budg_shel_home_ins = "" Then snap_budg_shel_home_ins = 0
			If snap_budg_shel_electricity = "" Then snap_budg_shel_electricity = 0
			If snap_budg_shel_heat_ac = "" Then snap_budg_shel_heat_ac = 0
			If snap_budg_shel_water_garbage = "" Then snap_budg_shel_water_garbage = 0
			If snap_budg_shel_phone = "" Then snap_budg_shel_phone = 0
			If snap_budg_shel_other = "" Then snap_budg_shel_other = 0

			snap_budg_shel_rent_mort = snap_budg_shel_rent_mort*1
			snap_budg_shel_prop_tax = snap_budg_shel_prop_tax*1
			snap_budg_shel_home_ins = snap_budg_shel_home_ins*1
			snap_budg_shel_electricity = snap_budg_shel_electricity*1
			snap_budg_shel_heat_ac = snap_budg_shel_heat_ac*1
			snap_budg_shel_water_garbage = snap_budg_shel_water_garbage*1
			snap_budg_shel_phone = snap_budg_shel_phone*1
			snap_budg_shel_other = snap_budg_shel_other*1

			snap_budg_housing_exp_total = snap_budg_shel_rent_mort + snap_budg_shel_prop_tax + snap_budg_shel_home_ins + snap_budg_shel_other
			snap_budg_utilities_exp_total = snap_budg_shel_electricity + snap_budg_shel_heat_ac + snap_budg_shel_phone

			snap_budg_utilities_list = "None"
			If snap_budg_shel_heat_ac <> 0 Then
				snap_budg_utilities_list = "Heat and AC"
			ElseIf snap_budg_shel_electricity <> 0 and snap_budg_shel_phone <> 0 Then
				snap_budg_utilities_list = "Electricity and Phone"
			ElseIf snap_budg_shel_electricity <> 0 Then
				snap_budg_utilities_list = "Electricity"
			ElseIf snap_budg_shel_phone <> 0 Then
				snap_budg_utilities_list = "Phone"
			End If
			snap_budg_housing_exp_total = FormatNumber(snap_budg_housing_exp_total, 2, -1, 0, -1)
			snap_budg_utilities_exp_total = FormatNumber(snap_budg_utilities_exp_total, 2, -1, 0, -1)

			snap_budg_shel_rent_mort = FormatNumber(snap_budg_shel_rent_mort, 2, -1, 0, -1)
			snap_budg_shel_prop_tax = FormatNumber(snap_budg_shel_prop_tax, 2, -1, 0, -1)
			snap_budg_shel_home_ins = FormatNumber(snap_budg_shel_home_ins, 2, -1, 0, -1)
			snap_budg_shel_electricity = FormatNumber(snap_budg_shel_electricity, 2, -1, 0, -1)
			snap_budg_shel_heat_ac = FormatNumber(snap_budg_shel_heat_ac, 2, -1, 0, -1)
			snap_budg_shel_water_garbage = FormatNumber(snap_budg_shel_water_garbage, 2, -1, 0, -1)
			snap_budg_shel_phone = FormatNumber(snap_budg_shel_phone, 2, -1, 0, -1)
			snap_budg_shel_other = FormatNumber(snap_budg_shel_other, 2, -1, 0, -1)

			EMReadScreen snap_budg_max_allow_shel, 			10, 5, 71
			EMReadScreen snap_budg_shel_expenses, 			10, 6, 71
			' EMReadScreen fsb2_net_adj_inc, 				10, 7, 71
			EMReadScreen snap_budg_max_net_adj_inc, 		10, 8, 71
			EMReadScreen snap_benefit_monthly_fs_allot, 	10, 10, 71
			EMReadScreen snap_benefit_drug_felon_sanc_amt, 	10, 12, 71
			EMReadScreen snap_benefit_amt_already_issued, 	 10, 13, 71
			EMReadScreen snap_benefit_recoup_amount, 		10, 14, 71
			EMReadScreen snap_benefit_benefit_amount, 		10, 16, 71
			EMReadScreen snap_benefit_state_food_amt, 		10, 17, 71
			EMReadScreen snap_benefit_fed_food_amt, 		10, 18, 71

			snap_budg_max_allow_shel = trim(snap_budg_max_allow_shel)
			snap_budg_shel_expenses = trim(snap_budg_shel_expenses)
			' fsb2_net_adj_inc = trim(fsb2_net_adj_inc)
			snap_budg_max_net_adj_inc = trim(snap_budg_max_net_adj_inc)
			snap_benefit_monthly_fs_allot = trim(snap_benefit_monthly_fs_allot)
			If snap_benefit_monthly_fs_allot = "" Then snap_benefit_monthly_fs_allot = "0.00"
			snap_benefit_drug_felon_sanc_amt = trim(snap_benefit_drug_felon_sanc_amt)
			snap_benefit_amt_already_issued = trim(snap_benefit_amt_already_issued)
			snap_benefit_recoup_amount = trim(snap_benefit_recoup_amount)
			snap_benefit_benefit_amount = trim(snap_benefit_benefit_amount)
			snap_benefit_state_food_amt = trim(snap_benefit_state_food_amt)
			snap_benefit_fed_food_amt = trim(snap_benefit_fed_food_amt)


			Call write_value_and_transmit("X", 14, 4)		''Resource Detail
			row = 8
			Do
				EMReadScreen ref_numb, 2, row, 12

				For case_memb = 0 to UBound(snap_elig_ref_numbs)
					If ref_numb = snap_elig_ref_numbs(case_memb) Then
						EMReadScreen memb_drug_felon_test, 6, row, 64
						snap_elig_membs_drug_felon_test(case_memb) = trim(memb_drug_felon_test)
					End If
				Next

				row = row + 1
				EMReadScreen next_ref_numb, 2, row, 12
			Loop until next_ref_numb = "  "
			transmit

			Call write_value_and_transmit("X", 14, 4)		''Resource Detail
			EMReadScreen snap_benefit_recoup_from_fed_fs, 10, 5, 51
			EMReadScreen snap_benefit_recoup_from_state_fs, 10, 7, 51

			snap_benefit_recoup_from_fed_fs = trim(snap_benefit_recoup_from_fed_fs)
			snap_benefit_recoup_from_state_fs = trim(snap_benefit_recoup_from_state_fs)

			transmit

			transmit 		'FSSM
			EMReadScreen snap_approved_date, 			8, 3, 14
			EMReadScreen snap_date_last_approval, 		8, 5, 31
			EMReadScreen snap_curr_prog_status, 		10, 6, 31
			EMReadScreen snap_elig_result, 				10, 7, 31
			EMReadScreen snap_reporting_status, 		12, 8, 31
			EMReadScreen snap_info_source, 				4, 9, 31
			EMReadScreen snap_benefit, 					12, 10, 31
			EMReadScreen snap_elig_revw_date, 			8, 11, 31
			EMReadScreen snap_budget_cycle, 			5, 12, 31
			EMReadScreen snap_budg_numb_in_assist_unit, 2, 13, 31

			EMReadScreen snap_budg_total_resources, 		10, 5, 71
			EMReadScreen snap_budg_max_resources, 			10, 6, 71
			EMReadScreen snap_budg_net_adj_inc, 			10, 7, 71
			EMReadScreen snap_benefit_monthly_fs_allotment, 10, 8, 71
			EMReadScreen snap_benefit_prorated_amt, 		10, 9, 71
			EMReadScreen snap_benefit_prorated_date,		8, 9, 58
			EMReadScreen snap_benefit_amt, 					10, 13, 71

			snap_approved_date = trim(snap_approved_date)
			snap_date_last_approval = trim(snap_date_last_approval)
			snap_curr_prog_status = trim(snap_curr_prog_status)
			snap_elig_result = trim(snap_elig_result)
			snap_reporting_status = trim(snap_reporting_status)
			snap_info_source = trim(snap_info_source)
			snap_benefit = trim(snap_benefit)
			snap_elig_revw_date = trim(snap_elig_revw_date)
			snap_budget_cycle = trim(snap_budget_cycle)
			snap_budg_numb_in_assist_unit = trim(snap_budg_numb_in_assist_unit)
			snap_budg_total_resources = trim(snap_budg_total_resources)
			snap_budg_max_resources = trim(snap_budg_max_resources)
			snap_budg_net_adj_inc = trim(snap_budg_net_adj_inc)
			snap_benefit_monthly_fs_allotment = trim(snap_benefit_monthly_fs_allotment)
			snap_benefit_prorated_amt = trim(snap_benefit_prorated_amt)
			snap_benefit_prorated_date = trim(snap_benefit_prorated_date)
			snap_benefit_amt = trim(snap_benefit_amt)

			If snap_budg_net_adj_inc = "" Then snap_budg_net_adj_inc = 0
			snap_budg_net_adj_inc = snap_budg_net_adj_inc*1
			snap_bug_30_percent_net_adj_inc = .3 * snap_budg_net_adj_inc
			snap_bug_30_percent_net_adj_inc = Round(snap_bug_30_percent_net_adj_inc)
			snap_budg_net_adj_inc = FormatNumber(snap_budg_net_adj_inc, 2, -1, 0, -1)
			snap_bug_30_percent_net_adj_inc = FormatNumber(snap_bug_30_percent_net_adj_inc, 2, -1, 0, -1)

			If snap_budg_numb_in_assist_unit = "" Then snap_budg_numb_in_assist_unit = 0
			snap_budg_numb_in_assist_unit = snap_budg_numb_in_assist_unit*1
			If snap_budg_numb_in_assist_unit = 0 Then snap_budg_thrifty_food_plan = "0.00"
			If snap_budg_numb_in_assist_unit = 1 Then snap_budg_thrifty_food_plan = "250.00"
			If snap_budg_numb_in_assist_unit = 2 Then snap_budg_thrifty_food_plan = "459.00"
			If snap_budg_numb_in_assist_unit = 3 Then snap_budg_thrifty_food_plan = "658.00"
			If snap_budg_numb_in_assist_unit = 4 Then snap_budg_thrifty_food_plan = "835.00"
			If snap_budg_numb_in_assist_unit = 5 Then snap_budg_thrifty_food_plan = "992.00"
			If snap_budg_numb_in_assist_unit = 6 Then snap_budg_thrifty_food_plan = "1,190.00"
			If snap_budg_numb_in_assist_unit = 7 Then snap_budg_thrifty_food_plan = "1,316.00"
			If snap_budg_numb_in_assist_unit = 8 Then snap_budg_thrifty_food_plan = "1,504.00"
			If snap_budg_numb_in_assist_unit > 8 Then snap_budg_thrifty_food_plan = 1504 + ((snap_budg_numb_in_assist_unit-8)*188)
			snap_budg_thrifty_food_plan = snap_budg_thrifty_food_plan & ""

			EMReadScreen fssm_expedited_info_exists, 16, 14, 44
			If fssm_expedited_info_exists = "EXPEDITED STATUS" Then
				Call write_value_and_transmit("X", 14, 72)		''Resource Detail
				EMReadScreen exp_status_issuance_on_or_before_15th, 1, 3, 5
				EMReadScreen exp_status_issuance_after_15th, 1, 5, 5
				EMReadScreen exp_status_issuance_app_month_fs_denial, 1, 9, 5

				EMReadScreen snap_exp_criteria_migrant_destitute, 1, 15, 5
				EMReadScreen snap_exp_criteria_resource_100_income_150, 1, 16, 5
				EMReadScreen snap_exp_criteria_resource_income_less_shelter, 1, 19, 5

				EMReadScreen snap_exp_verif_status_postponed, 1, 15, 52
				EMReadScreen snap_exp_verif_status_out_of_state, 1, 17, 52
				EMReadScreen snap_exp_verif_status_all_provided, 1, 19, 52
				transmit

				If exp_status_issuance_on_or_before_15th = "X" Then snap_exp_package_includes_month_one = True
				If exp_status_issuance_after_15th = "X" Then
					snap_exp_package_includes_month_one = True
					snap_exp_package_includes_month_two = True
				End If
				If exp_status_issuance_app_month_fs_denial = "X" Then snap_exp_package_includes_month_two = True

				If snap_exp_criteria_migrant_destitute = "X" Then snap_exp_criteria_migrant_destitute = True
				If snap_exp_criteria_migrant_destitute = "_" Then snap_exp_criteria_migrant_destitute = False
				If snap_exp_criteria_resource_100_income_150 = "X" Then snap_exp_criteria_resource_100_income_150 = True
				If snap_exp_criteria_resource_100_income_150 = "_" Then snap_exp_criteria_resource_100_income_150 = False
				If snap_exp_criteria_resource_income_less_shelter = "X" Then snap_exp_criteria_resource_income_less_shelter = True
				If snap_exp_criteria_resource_income_less_shelter = "_" Then snap_exp_criteria_resource_income_less_shelter = False

				If snap_exp_verif_status_postponed = "X" Then snap_exp_verif_status_postponed = True
				If snap_exp_verif_status_postponed = "_" Then snap_exp_verif_status_postponed = False
				If snap_exp_verif_status_out_of_state = "X" Then snap_exp_verif_status_out_of_state = True
				If snap_exp_verif_status_out_of_state = "_" Then snap_exp_verif_status_out_of_state = False
				If snap_exp_verif_status_all_provided = "X" Then snap_exp_verif_status_all_provided = True
				If snap_exp_verif_status_all_provided = "_" Then snap_exp_verif_status_all_provided = False


			End If

			EMReadScreen snap_elig_worker_message_one, 80, 17, 1
			EMReadScreen snap_elig_worker_message_two, 80, 18, 1

			snap_elig_worker_message_one = trim(snap_elig_worker_message_one)
			snap_elig_worker_message_two = trim(snap_elig_worker_message_two)

			If snap_budg_total_earned_inc <> "" Then snap_earned_income_budgeted = True
			If snap_budg_total_unea_inc <> "" Then snap_unearned_income_budgeted = True
			If snap_budg_shel_rent_mort <> "" or snap_budg_shel_prop_tax <> "" or snap_budg_shel_home_ins <> "" or snap_budg_shel_other <> ""Then snap_shel_costs_budgeted = True
			If snap_budg_shel_electricity <> "" or snap_budg_shel_heat_ac <> "" or snap_budg_shel_water_garbage <> "" or snap_budg_shel_phone <> ""Then snap_hest_costs_budgeted = True
			' categorical_eligibility = ""
		End If

		Call Back_to_SELF
	End sub
end class

class hc_eligibility_detail
	public elig_footer_month
	public elig_footer_year
	' public elig_version_number
	' public elig_version_date
	' public elig_version_result
	public approved_today
	public approved_version_found
	public er_month
	public hrf_month
	public er_status
	public er_caf_date
	public er_interview_date
	public hrf_status
	public hrf_doc_date

	public hc_elig_ref_numbs()
	public hc_elig_full_name()
	public hc_prog_elig_appd()
	public hc_prog_elig_major_program()
	public hc_prog_elig_eligibility_result()
	public hc_prog_elig_status()
	public hc_prog_elig_app_indc()
	public hc_prog_elig_magi_excempt()
	public hc_prog_elig_process_date()
	public hc_prog_elig_source_of_info()
	public hc_prog_elig_responsible_county()
	public hc_prog_elig_servicing_county()
	public hc_prog_elig_test_application_withdrawn()
	public hc_prog_elig_test_application_process_incomplete()
	public hc_prog_elig_test_no_new_prog_eligibility()
	public hc_prog_elig_test_assistance_unit()
	public hc_prog_elig_worker_msg_one()
	public hc_prog_elig_worker_msg_two()
	public hc_prog_elig_elig_type()
	public hc_prog_elig_elig_standard()
	public hc_prog_elig_method()
	public hc_prog_elig_waiver()
	public hc_prog_elig_total_net_income()
	public hc_prog_elig_standard()
	public hc_prog_elig_excess_income()
	public hc_prog_elig_test_absence()
	public hc_prog_elig_test_assets()
	public hc_prog_elig_test_citizenship()
	public hc_prog_elig_test_coop()
	public hc_prog_elig_test_correctional_faci()
	public hc_prog_elig_test_death()
	public hc_prog_elig_test_elig_other_prog()
	public hc_prog_elig_test_fail_file()
	public hc_prog_elig_test_IMD()
	public hc_prog_elig_test_uncompensated_transfer()
	public hc_prog_elig_test_income()
	public hc_prog_elig_test_medicare_elig()
	public hc_prog_elig_test_MNSure_system()
	public hc_prog_elig_test_Obligation_one_mo()
	public hc_prog_elig_test_obligation_six_mo()
	public hc_prog_elig_test_other_health_ins()
	public hc_prog_elig_test_parent()
	public hc_prog_elig_test_residence()
	public hc_prog_elig_test_verif()
	public hc_prog_elig_test_withdrawn()
	public hc_prog_elig_test_coop_pben_cash()
	public hc_prog_elig_test_coop_pben_smrt()
	public hc_prog_elig_test_coop_pben()
	public hc_prog_elig_test_coop_fail_provide_info()
	public hc_prog_elig_test_coop_IEVS()
	public hc_prog_elig_test_coop_medical_support()
	public hc_prog_elig_test_coop_other_health_ins()
	public hc_prog_elig_test_coop_SSN()
	public hc_prog_elig_test_coop_third_party_liability()
	public hc_prog_elig_test_fail_file_HRF()
	public hc_prog_elig_test_fail_file_IR()
	public hc_prog_elig_test_fail_file_AR()
	public hc_prog_elig_test_fail_file_ER()
	public hc_prog_elig_test_fail_file_quarterly_TYMA()
	public hc_prog_elig_test_verif_ACCT()
	public hc_prog_elig_test_verif_BUSI()
	public hc_prog_elig_test_verif_JOBS()
	public hc_prog_elig_test_verif_IMIG_status()
	public hc_prog_elig_test_verif_LUMP()
	public hc_prog_elig_test_verif_OTHR()
	public hc_prog_elig_test_verif_PBEN()
	public hc_prog_elig_test_verif_PREG()
	public hc_prog_elig_test_verif_RBIC()
	public hc_prog_elig_test_verif_REST()
	public hc_prog_elig_test_verif_SECU()
	public hc_prog_elig_test_verif_SPON()
	public hc_prog_elig_test_verif_TRAN()
	public hc_prog_elig_test_verif_UNEA()
	public hc_prog_elig_test_verif_cit_id()
	public hc_prog_elig_test_verif_CARS()
	public hc_prog_elig_hh_size()
	public hc_prog_elig_members_whose_income_counts()
	public hc_prog_elig_members_whose_income_counts_list()
	public hc_prog_elig_PTMA()
	public hc_prog_elig_elig_standard_percent()
	public hc_prog_elig_basis()
	public hc_prog_elig_budg_gross_unearned()
	public hc_prog_elig_budg_excluded_unearned()
	public hc_prog_elig_budg_unearned_deduction()
	public hc_prog_elig_budg_net_unearned_income()
	public hc_prog_elig_budg_gross_earned()
	public hc_prog_elig_budg_excluded_earned()
	public hc_prog_elig_budg_work_exp_deduction()
	public hc_prog_elig_budg_earned_disregarrd()
	public hc_prog_elig_budg_dependent_care()
	public hc_prog_elig_budg_earned_deduction()
	public hc_prog_elig_budg_net_earned_income()
	public hc_prog_elig_budg_child_sup_deduction()
	public hc_prog_elig_budg_deemed_income()
	public hc_prog_elig_budg_total_net_income()
	public hc_prog_elig_budg_income_standard()
	public hc_prog_elig_budg_spenddown()
	public hc_prog_elig_budg_transfer_penalty()
	public hc_prog_elig_budg_total_liability()
	public hc_prog_elig_budg_deemed_earned()
	public hc_prog_elig_budg_countable_earrned_income()
	public hc_prog_elig_budg_subtotal_countable_income()
	public hc_prog_elig_budg_va_aid_attendincome()
	public hc_prog_elig_budg_total_countable_income()
	public hc_prog_elig_budg_LTC_exclusions()
	public hc_prog_elig_budg_medicare_premium()
	public hc_prog_elig_budg_maint_needs_allowance()
	public hc_prog_elig_budg_guardian_rep_payee_fee()
	public hc_prog_elig_budg_spousal_allocation()
	public hc_prog_elig_budg_family_allocation()
	public hc_prog_elig_budg_health_ins_premium()
	public hc_prog_elig_budg_other_medical_expense()
	public hc_prog_elig_budg_SSI_1611_benefit()
	public hc_prog_elig_budg_other_deductions()
	public hc_prog_elig_budg_waiver_obligation()
	public hc_prog_elig_budg_person_clothing_needs()
	public hc_prog_elig_budg_LTC_spenddown()
	public hc_prog_elig_budg_medical_spenddown()
	public hc_prog_elig_mobl_result()
	public hc_prog_elig_mobl_type()
	public hc_prog_elig_mobl_period()
	public hc_prog_elig_spdn_option()
	public hc_prog_elig_spdn_type()
	public hc_prog_elig_spdn_method()
	public hc_prog_elig_spdn_covered_pop()
	public hc_prog_elig_original_monthly_spdn()
	public hc_prog_elig_monthly_spdn_counted_bills()
	public hc_prog_elig_monthly_spdn_satisfaction_date()
	public hc_prog_elig_monthly_spdn_recipient_amount()
	public hc_prog_elig_monthly_spdn_balance()
	public hc_prog_elig_oiginal_six_mo_spdn()
	public hc_prog_elig_six_mo_spdn_counted_bills()
	public hc_prog_elig_six_mo_spnd_satisfaction_date()
	public hc_prog_elig_six_mo_spdn_recipient_amount()
	public hc_prog_elig_six_mo_spdn_unused_balance()
	public hc_prog_elig_cert_prd_total_net_income()
	public hc_prog_elig_cert_prd_total_standard()
	public hc_prog_elig_cert_prd_total_excess_income()
	public hc_prog_elig_app_date()
	public hc_prog_elig_curr_prog_status()
	public hc_prog_elig_elig_result()
	public hc_prog_elig_elig_begin_date()
	public hc_prog_elig_HRF_reporting()
	public hc_prog_elig_ER_date()
	public hc_prog_elig_SR_date()
	public hc_prog_elig_TYMA_begin_date()
	public hc_prog_elig_TYMA_type()
	public hc_prog_elig_budg_deemed_unearned()
	public hc_prog_elig_budg_standard_disregard()
	public hc_prog_elig_budg_excess_income()
	public hc_prog_elig_test_after_processing_month()
	public hc_prog_elig_test_basis_for_other_prog()
	public hc_prog_elig_test_medicare_part_a()
	' public

	public sub read_elig()
		approved_today = False
		approved_version_found = False

		ReDim hc_elig_ref_numbs(0)
		ReDim hc_elig_full_name(0)
		ReDim hc_prog_elig_appd(0)
		ReDim hc_prog_elig_major_program(0)
		ReDim hc_prog_elig_eligibility_result(0)
		ReDim hc_prog_elig_status(0)
		ReDim hc_prog_elig_app_indc(0)
		ReDim hc_prog_elig_magi_excempt(0)
		ReDim hc_prog_elig_process_date(0)
		ReDim hc_prog_elig_source_of_info(0)
		ReDim hc_prog_elig_responsible_county(0)
		ReDim hc_prog_elig_servicing_county(0)
		ReDim hc_prog_elig_test_application_withdrawn(0)
		ReDim hc_prog_elig_test_application_process_incomplete(0)
		ReDim hc_prog_elig_test_no_new_prog_eligibility(0)
		ReDim hc_prog_elig_test_assistance_unit(0)
		ReDim hc_prog_elig_worker_msg_one(0)
		ReDim hc_prog_elig_worker_msg_two(0)
		ReDim hc_prog_elig_elig_type(0)
		ReDim hc_prog_elig_elig_standard(0)
		ReDim hc_prog_elig_method(0)
		ReDim hc_prog_elig_waiver(0)
		ReDim hc_prog_elig_total_net_income(0)
		ReDim hc_prog_elig_standard(0)
		ReDim hc_prog_elig_excess_income(0)
		ReDim hc_prog_elig_test_absence(0)
		ReDim hc_prog_elig_test_assets(0)
		ReDim hc_prog_elig_test_citizenship(0)
		ReDim hc_prog_elig_test_coop(0)
		ReDim hc_prog_elig_test_correctional_faci(0)
		ReDim hc_prog_elig_test_death(0)
		ReDim hc_prog_elig_test_elig_other_prog(0)
		ReDim hc_prog_elig_test_fail_file(0)
		ReDim hc_prog_elig_test_IMD(0)
		ReDim hc_prog_elig_test_uncompensated_transfer(0)
		ReDim hc_prog_elig_test_income(0)
		ReDim hc_prog_elig_test_medicare_elig(0)
		ReDim hc_prog_elig_test_MNSure_system(0)
		ReDim hc_prog_elig_test_Obligation_one_mo(0)
		ReDim hc_prog_elig_test_obligation_six_mo(0)
		ReDim hc_prog_elig_test_other_health_ins(0)
		ReDim hc_prog_elig_test_parent(0)
		ReDim hc_prog_elig_test_residence(0)
		ReDim hc_prog_elig_test_verif(0)
		ReDim hc_prog_elig_test_withdrawn(0)
		ReDim hc_prog_elig_test_coop_pben_cash(0)
		ReDim hc_prog_elig_test_coop_pben_smrt(0)
		ReDim hc_prog_elig_test_coop_pben(0)
		ReDim hc_prog_elig_test_coop_fail_provide_info(0)
		ReDim hc_prog_elig_test_coop_IEVS(0)
		ReDim hc_prog_elig_test_coop_medical_support(0)
		ReDim hc_prog_elig_test_coop_other_health_ins(0)
		ReDim hc_prog_elig_test_coop_SSN(0)
		ReDim hc_prog_elig_test_coop_third_party_liability(0)
		ReDim hc_prog_elig_test_fail_file_HRF(0)
		ReDim hc_prog_elig_test_fail_file_IR(0)
		ReDim hc_prog_elig_test_fail_file_AR(0)
		ReDim hc_prog_elig_test_fail_file_ER(0)
		ReDim hc_prog_elig_test_fail_file_quarterly_TYMA(0)
		ReDim hc_prog_elig_test_verif_ACCT(0)
		ReDim hc_prog_elig_test_verif_BUSI(0)
		ReDim hc_prog_elig_test_verif_JOBS(0)
		ReDim hc_prog_elig_test_verif_IMIG_status(0)
		ReDim hc_prog_elig_test_verif_LUMP(0)
		ReDim hc_prog_elig_test_verif_OTHR(0)
		ReDim hc_prog_elig_test_verif_PBEN(0)
		ReDim hc_prog_elig_test_verif_PREG(0)
		ReDim hc_prog_elig_test_verif_RBIC(0)
		ReDim hc_prog_elig_test_verif_REST(0)
		ReDim hc_prog_elig_test_verif_SECU(0)
		ReDim hc_prog_elig_test_verif_SPON(0)
		ReDim hc_prog_elig_test_verif_TRAN(0)
		ReDim hc_prog_elig_test_verif_UNEA(0)
		ReDim hc_prog_elig_test_verif_cit_id(0)
		ReDim hc_prog_elig_test_verif_CARS(0)
		ReDim hc_prog_elig_hh_size(0)
		ReDim hc_prog_elig_members_whose_income_counts(0)
		ReDim hc_prog_elig_members_whose_income_counts_list(0)
		ReDim hc_prog_elig_PTMA(0)
		ReDim hc_prog_elig_elig_standard_percent(0)
		ReDim hc_prog_elig_basis(0)
		ReDim hc_prog_elig_budg_gross_unearned(0)
		ReDim hc_prog_elig_budg_excluded_unearned(0)
		ReDim hc_prog_elig_budg_unearned_deduction(0)
		ReDim hc_prog_elig_budg_net_unearned_income(0)
		ReDim hc_prog_elig_budg_gross_earned(0)
		ReDim hc_prog_elig_budg_excluded_earned(0)
		ReDim hc_prog_elig_budg_work_exp_deduction(0)
		ReDim hc_prog_elig_budg_earned_disregarrd(0)
		ReDim hc_prog_elig_budg_dependent_care(0)
		ReDim hc_prog_elig_budg_earned_deduction(0)
		ReDim hc_prog_elig_budg_net_earned_income(0)
		ReDim hc_prog_elig_budg_child_sup_deduction(0)
		ReDim hc_prog_elig_budg_deemed_income(0)
		ReDim hc_prog_elig_budg_total_net_income(0)
		ReDim hc_prog_elig_budg_income_standard(0)
		ReDim hc_prog_elig_budg_spenddown(0)
		ReDim hc_prog_elig_budg_transfer_penalty(0)
		ReDim hc_prog_elig_budg_total_liability(0)
		ReDim hc_prog_elig_budg_deemed_earned(0)
		ReDim hc_prog_elig_budg_countable_earrned_income(0)
		ReDim hc_prog_elig_budg_subtotal_countable_income(0)
		ReDim hc_prog_elig_budg_va_aid_attendincome(0)
		ReDim hc_prog_elig_budg_total_countable_income(0)
		ReDim hc_prog_elig_budg_LTC_exclusions(0)
		ReDim hc_prog_elig_budg_medicare_premium(0)
		ReDim hc_prog_elig_budg_maint_needs_allowance(0)
		ReDim hc_prog_elig_budg_guardian_rep_payee_fee(0)
		ReDim hc_prog_elig_budg_spousal_allocation(0)
		ReDim hc_prog_elig_budg_family_allocation(0)
		ReDim hc_prog_elig_budg_health_ins_premium(0)
		ReDim hc_prog_elig_budg_other_medical_expense(0)
		ReDim hc_prog_elig_budg_SSI_1611_benefit(0)
		ReDim hc_prog_elig_budg_other_deductions(0)
		ReDim hc_prog_elig_budg_waiver_obligation(0)
		ReDim hc_prog_elig_budg_person_clothing_needs(0)
		ReDim hc_prog_elig_budg_LTC_spenddown(0)
		ReDim hc_prog_elig_budg_medical_spenddown(0)
		ReDim hc_prog_elig_mobl_result(0)
		ReDim hc_prog_elig_mobl_type(0)
		ReDim hc_prog_elig_mobl_period(0)
		ReDim hc_prog_elig_spdn_option(0)
		ReDim hc_prog_elig_spdn_type(0)
		ReDim hc_prog_elig_spdn_method(0)
		ReDim hc_prog_elig_spdn_covered_pop(0)
		ReDim hc_prog_elig_original_monthly_spdn(0)
		ReDim hc_prog_elig_monthly_spdn_counted_bills(0)
		ReDim hc_prog_elig_monthly_spdn_satisfaction_date(0)
		ReDim hc_prog_elig_monthly_spdn_recipient_amount(0)
		ReDim hc_prog_elig_monthly_spdn_balance(0)
		ReDim hc_prog_elig_oiginal_six_mo_spdn(0)
		ReDim hc_prog_elig_six_mo_spdn_counted_bills(0)
		ReDim hc_prog_elig_six_mo_spnd_satisfaction_date(0)
		ReDim hc_prog_elig_six_mo_spdn_recipient_amount(0)
		ReDim hc_prog_elig_six_mo_spdn_unused_balance(0)
		ReDim hc_prog_elig_cert_prd_total_net_income(0)
		ReDim hc_prog_elig_cert_prd_total_standard(0)
		ReDim hc_prog_elig_cert_prd_total_excess_income(0)
		ReDim hc_prog_elig_app_date(0)
		ReDim hc_prog_elig_curr_prog_status(0)
		ReDim hc_prog_elig_elig_result(0)
		ReDim hc_prog_elig_elig_begin_date(0)
		ReDim hc_prog_elig_HRF_reporting(0)
		ReDim hc_prog_elig_ER_date(0)
		ReDim hc_prog_elig_SR_date(0)
		ReDim hc_prog_elig_TYMA_begin_date(0)
		ReDim hc_prog_elig_TYMA_type(0)
		ReDim hc_prog_elig_budg_deemed_unearned(0)
		ReDim hc_prog_elig_budg_standard_disregard(0)
		ReDim hc_prog_elig_budg_excess_income(0)
		ReDim hc_prog_elig_test_after_processing_month(0)
		ReDim hc_prog_elig_test_basis_for_other_prog(0)
		ReDim hc_prog_elig_test_medicare_part_a(0)

		call navigate_to_MAXIS_screen("ELIG", "HC  ")
		EMWriteScreen elig_footer_month, 19, 54
		EMWriteScreen elig_footer_year, 19, 57
		transmit

		hc_row = 8
		hc_prog_count = 0
		Do
			ReDim preserve hc_elig_ref_numbs(hc_prog_count)
			ReDim preserve hc_elig_full_name(hc_prog_count)
			ReDim preserve hc_prog_elig_appd(hc_prog_count)
			ReDim preserve hc_prog_elig_major_program(hc_prog_count)
			ReDim preserve hc_prog_elig_eligibility_result(hc_prog_count)
			ReDim preserve hc_prog_elig_status(hc_prog_count)
			ReDim preserve hc_prog_elig_app_indc(hc_prog_count)
			ReDim preserve hc_prog_elig_magi_excempt(hc_prog_count)
			ReDim preserve hc_prog_elig_process_date(hc_prog_count)
			ReDim preserve hc_prog_elig_source_of_info(hc_prog_count)
			ReDim preserve hc_prog_elig_responsible_county(hc_prog_count)
			ReDim preserve hc_prog_elig_servicing_county(hc_prog_count)
			ReDim preserve hc_prog_elig_test_application_withdrawn(hc_prog_count)
			ReDim preserve hc_prog_elig_test_application_process_incomplete(hc_prog_count)
			ReDim preserve hc_prog_elig_test_no_new_prog_eligibility(hc_prog_count)
			ReDim preserve hc_prog_elig_test_assistance_unit(hc_prog_count)
			ReDim preserve hc_prog_elig_worker_msg_one(hc_prog_count)
			ReDim preserve hc_prog_elig_worker_msg_two(hc_prog_count)
			ReDim preserve hc_prog_elig_elig_type(hc_prog_count)
			ReDim preserve hc_prog_elig_elig_standard(hc_prog_count)
			ReDim preserve hc_prog_elig_method(hc_prog_count)
			ReDim preserve hc_prog_elig_waiver(hc_prog_count)
			ReDim preserve hc_prog_elig_total_net_income(hc_prog_count)
			ReDim preserve hc_prog_elig_standard(hc_prog_count)
			ReDim preserve hc_prog_elig_excess_income(hc_prog_count)
			ReDim preserve hc_prog_elig_test_absence(hc_prog_count)
			ReDim preserve hc_prog_elig_test_assets(hc_prog_count)
			ReDim preserve hc_prog_elig_test_citizenship(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop(hc_prog_count)
			ReDim preserve hc_prog_elig_test_correctional_faci(hc_prog_count)
			ReDim preserve hc_prog_elig_test_death(hc_prog_count)
			ReDim preserve hc_prog_elig_test_elig_other_prog(hc_prog_count)
			ReDim preserve hc_prog_elig_test_fail_file(hc_prog_count)
			ReDim preserve hc_prog_elig_test_IMD(hc_prog_count)
			ReDim preserve hc_prog_elig_test_uncompensated_transfer(hc_prog_count)
			ReDim preserve hc_prog_elig_test_income(hc_prog_count)
			ReDim preserve hc_prog_elig_test_medicare_elig(hc_prog_count)
			ReDim preserve hc_prog_elig_test_MNSure_system(hc_prog_count)
			ReDim preserve hc_prog_elig_test_Obligation_one_mo(hc_prog_count)
			ReDim preserve hc_prog_elig_test_obligation_six_mo(hc_prog_count)
			ReDim preserve hc_prog_elig_test_other_health_ins(hc_prog_count)
			ReDim preserve hc_prog_elig_test_parent(hc_prog_count)
			ReDim preserve hc_prog_elig_test_residence(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif(hc_prog_count)
			ReDim preserve hc_prog_elig_test_withdrawn(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_pben_cash(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_pben_smrt(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_pben(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_fail_provide_info(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_IEVS(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_medical_support(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_other_health_ins(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_SSN(hc_prog_count)
			ReDim preserve hc_prog_elig_test_coop_third_party_liability(hc_prog_count)
			ReDim preserve hc_prog_elig_test_fail_file_HRF(hc_prog_count)
			ReDim preserve hc_prog_elig_test_fail_file_IR(hc_prog_count)
			ReDim preserve hc_prog_elig_test_fail_file_AR(hc_prog_count)
			ReDim preserve hc_prog_elig_test_fail_file_ER(hc_prog_count)
			ReDim preserve hc_prog_elig_test_fail_file_quarterly_TYMA(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_ACCT(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_BUSI(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_JOBS(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_IMIG_status(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_LUMP(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_OTHR(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_PBEN(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_PREG(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_RBIC(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_REST(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_SECU(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_SPON(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_TRAN(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_UNEA(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_cit_id(hc_prog_count)
			ReDim preserve hc_prog_elig_test_verif_CARS(hc_prog_count)
			ReDim preserve hc_prog_elig_hh_size(hc_prog_count)
			ReDim preserve hc_prog_elig_members_whose_income_counts(hc_prog_count)
			ReDim preserve hc_prog_elig_members_whose_income_counts_list(hc_prog_count)
			ReDim preserve hc_prog_elig_PTMA(hc_prog_count)
			ReDim preserve hc_prog_elig_elig_standard_percent(hc_prog_count)
			ReDim preserve hc_prog_elig_basis(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_gross_unearned(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_excluded_unearned(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_unearned_deduction(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_net_unearned_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_gross_earned(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_excluded_earned(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_work_exp_deduction(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_earned_disregarrd(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_dependent_care(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_earned_deduction(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_net_earned_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_child_sup_deduction(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_deemed_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_total_net_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_income_standard(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_spenddown(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_transfer_penalty(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_total_liability(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_deemed_earned(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_countable_earrned_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_subtotal_countable_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_va_aid_attendincome(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_total_countable_income(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_LTC_exclusions(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_medicare_premium(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_maint_needs_allowance(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_guardian_rep_payee_fee(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_spousal_allocation(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_family_allocation(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_health_ins_premium(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_other_medical_expense(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_SSI_1611_benefit(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_other_deductions(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_waiver_obligation(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_person_clothing_needs(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_LTC_spenddown(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_medical_spenddown(hc_prog_count)
			ReDim preserve hc_prog_elig_mobl_result(hc_prog_count)
			ReDim preserve hc_prog_elig_mobl_type(hc_prog_count)
			ReDim preserve hc_prog_elig_mobl_period(hc_prog_count)
			ReDim preserve hc_prog_elig_spdn_option(hc_prog_count)
			ReDim preserve hc_prog_elig_spdn_type(hc_prog_count)
			ReDim preserve hc_prog_elig_spdn_method(hc_prog_count)
			ReDim preserve hc_prog_elig_spdn_covered_pop(hc_prog_count)
			ReDim preserve hc_prog_elig_original_monthly_spdn(hc_prog_count)
			ReDim preserve hc_prog_elig_monthly_spdn_counted_bills(hc_prog_count)
			ReDim preserve hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count)
			ReDim preserve hc_prog_elig_monthly_spdn_recipient_amount(hc_prog_count)
			ReDim preserve hc_prog_elig_monthly_spdn_balance(hc_prog_count)
			ReDim preserve hc_prog_elig_oiginal_six_mo_spdn(hc_prog_count)
			ReDim preserve hc_prog_elig_six_mo_spdn_counted_bills(hc_prog_count)
			ReDim preserve hc_prog_elig_six_mo_spnd_satisfaction_date(hc_prog_count)
			ReDim preserve hc_prog_elig_six_mo_spdn_recipient_amount(hc_prog_count)
			ReDim preserve hc_prog_elig_six_mo_spdn_unused_balance(hc_prog_count)
			ReDim preserve hc_prog_elig_cert_prd_total_net_income(hc_prog_count)
			ReDim preserve hc_prog_elig_cert_prd_total_standard(hc_prog_count)
			ReDim preserve hc_prog_elig_cert_prd_total_excess_income(hc_prog_count)
			ReDim preserve hc_prog_elig_app_date(hc_prog_count)
			ReDim preserve hc_prog_elig_curr_prog_status(hc_prog_count)
			ReDim preserve hc_prog_elig_elig_result(hc_prog_count)
			ReDim preserve hc_prog_elig_elig_begin_date(hc_prog_count)
			ReDim preserve hc_prog_elig_HRF_reporting(hc_prog_count)
			ReDim preserve hc_prog_elig_ER_date(hc_prog_count)
			ReDim preserve hc_prog_elig_SR_date(hc_prog_count)
			ReDim preserve hc_prog_elig_TYMA_begin_date(hc_prog_count)
			ReDim preserve hc_prog_elig_TYMA_type(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_deemed_unearned(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_standard_disregard(hc_prog_count)
			ReDim preserve hc_prog_elig_budg_excess_income(hc_prog_count)
			ReDim preserve hc_prog_elig_test_after_processing_month(hc_prog_count)
			ReDim preserve hc_prog_elig_test_basis_for_other_prog(hc_prog_count)
			ReDim preserve hc_prog_elig_test_medicare_part_a(hc_prog_count)

			EMReadScreen hc_elig_ref_numbs(hc_prog_count), 2, hc_row, 3
			EMReadScreen hc_elig_full_name(hc_prog_count), 17, hc_row, 7

			If hc_elig_ref_numbs(hc_prog_count) = "  " Then
				hc_elig_ref_numbs(hc_prog_count) = hc_elig_ref_numbs(hc_prog_count-1)
				hc_elig_full_name(hc_prog_count) = hc_elig_full_name(hc_prog_count-1)
			End If
			EMReadScreen clt_hc_prog, 4, hc_row, 28
			If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "    " Then

				EMReadScreen prog_status, 3, hc_row, 68
				If prog_status <> "APP" Then                        'Finding the approved version
					EMReadScreen total_versions, 2, hc_row, 64
					If total_versions = "01" Then
						hc_prog_elig_appd(hc_prog_count) = False
					Else
						EMReadScreen current_version, 2, hc_row, 58
						If current_version = "01" Then
							hc_prog_elig_appd(hc_prog_count) = False
						Else
							prev_version = right ("00" & abs(current_version) - 1, 2)
							EMWriteScreen prev_version, hc_row, 58
							transmit
							hc_prog_elig_appd(hc_prog_count) = True
						End If

					End If
				Else
					hc_prog_elig_appd(hc_prog_count) = True
				End If
			Else
				hc_prog_elig_appd(hc_prog_count) = False
			End If

			If hc_prog_elig_appd(hc_prog_count) = True Then
				EMReadScreen hc_prog_elig_major_program(hc_prog_count), 		4, hc_row, 28
				EMReadScreen hc_prog_elig_eligibility_result(hc_prog_count), 	8, hc_row, 41
				EMReadScreen hc_prog_elig_status(hc_prog_count), 				8, hc_row, 50
				EMReadScreen hc_prog_elig_app_indc(hc_prog_count), 				6, hc_row, 68
				EMReadScreen hc_prog_elig_magi_excempt(hc_prog_count), 			6, hc_row, 74

				hc_prog_elig_major_program(hc_prog_count) = trim(hc_prog_elig_major_program(hc_prog_count))

				Call write_value_and_transmit("X", hc_row, 26)
				' MsgBox "MOVING - 1" & vbCr & hc_prog_elig_major_program(hc_prog_count) & vbCr & "MEMB " & hc_elig_ref_numbs(hc_prog_count)
				EMReadScreen hc_prog_elig_process_date(hc_prog_count), 8, 2, 73
				hc_prog_elig_process_date(hc_prog_count) = DateAdd("d", 0, hc_prog_elig_process_date(hc_prog_count))

				' If DateDiff("'d", hc_prog_elig_process_date(hc_prog_count), date) = 0 Then
					If hc_prog_elig_major_program(hc_prog_count) = "HC D" Then
						EMReadScreen hc_prog_elig_source_of_info(hc_prog_count), 		4, 9, 33
						EMReadScreen hc_prog_elig_responsible_county(hc_prog_count), 	2, 8, 78
						EMReadScreen hc_prog_elig_servicing_county(hc_prog_count), 	2, 9, 78

						EMReadScreen hc_prog_elig_test_application_withdrawn(hc_prog_count), 			6, 13, 22
						EMReadScreen hc_prog_elig_test_application_process_incomplete(hc_prog_count), 6, 14, 22
						EMReadScreen hc_prog_elig_test_no_new_prog_eligibility(hc_prog_count), 		6, 15, 22
						EMReadScreen hc_prog_elig_test_assistance_unit(hc_prog_count), 				6, 16, 22

						EMReadScreen hc_prog_elig_worker_msg_one(hc_prog_count), 78, 19, 3
					End If

					If hc_prog_elig_major_program(hc_prog_count) = "MA" or hc_prog_elig_major_program(hc_prog_count) = "EMA" Then
						hc_col = 17
						Do
							EMReadScreen budg_mo, 2, 6, hc_col + 2
							EMReadScreen budg_yr, 2, 6, hc_col + 5
							' MsgBox "BUDG MO/YR:" & vbCr & budg_mo & "/" & budg_yr & vbCr & "Col: " & hc_col
							If budg_mo = elig_footer_month AND budg_yr = elig_footer_year Then
								EMReadScreen hc_prog_elig_elig_type(hc_prog_count), 		2, 12, hc_col
								EMReadScreen hc_prog_elig_elig_standard(hc_prog_count), 	1, 12, hc_col + 5
								EMReadScreen hc_prog_elig_method(hc_prog_count), 			1, 13, hc_col + 4
								EMReadScreen hc_prog_elig_waiver(hc_prog_count), 			1, 14, hc_col + 4

								EMReadScreen hc_prog_elig_total_net_income(hc_prog_count), 9, 15, hc_col
								EMReadScreen hc_prog_elig_standard(hc_prog_count), 		9, 16, hc_col
								EMReadScreen hc_prog_elig_excess_income(hc_prog_count), 	9, 17, hc_col

								Call write_value_and_transmit("X", 7, hc_col)						'Opening the HC Span
								' MsgBox "MOVING - 2"
								If hc_prog_elig_major_program(hc_prog_count) = "MA" or hc_prog_elig_major_program(hc_prog_count) = "EMA" Then
									EMReadScreen hc_prog_elig_test_absence(hc_prog_count), 			6, 6, 5
									EMReadScreen hc_prog_elig_test_assets(hc_prog_count), 			6, 7, 5
									EMReadScreen hc_prog_elig_test_assistance_unit(hc_prog_count), 	6, 8, 5
									EMReadScreen hc_prog_elig_test_citizenship(hc_prog_count), 		6, 9, 5
									EMReadScreen hc_prog_elig_test_coop(hc_prog_count), 				6, 10, 5
									EMReadScreen hc_prog_elig_test_correctional_faci(hc_prog_count), 	6, 11, 5
									EMReadScreen hc_prog_elig_test_death(hc_prog_count), 				6, 12, 5
									EMReadScreen hc_prog_elig_test_elig_other_prog(hc_prog_count), 	6, 13, 5
									EMReadScreen hc_prog_elig_test_fail_file(hc_prog_count), 			6, 14, 5
									EMReadScreen hc_prog_elig_test_IMD(hc_prog_count), 				6, 15, 5

									EMReadScreen hc_prog_elig_test_uncompensated_transfer(hc_prog_count), 6, 18, 5

									EMReadScreen hc_prog_elig_test_income(hc_prog_count), 			6, 6, 46
									EMReadScreen hc_prog_elig_test_medicare_elig(hc_prog_count), 		6, 7, 46
									EMReadScreen hc_prog_elig_test_MNSure_system(hc_prog_count), 		6, 8, 46
									EMReadScreen hc_prog_elig_test_Obligation_one_mo(hc_prog_count), 	6, 9, 46
									EMReadScreen hc_prog_elig_test_obligation_six_mo(hc_prog_count), 	6, 10, 46
									If hc_prog_elig_major_program(hc_prog_count) = "MA" Then
										EMReadScreen hc_prog_elig_test_other_health_ins(hc_prog_count), 6, 11, 46
										EMReadScreen hc_prog_elig_test_parent(hc_prog_count), 			6, 12, 46
										EMReadScreen hc_prog_elig_test_residence(hc_prog_count), 		6, 13, 46
										EMReadScreen hc_prog_elig_test_verif(hc_prog_count), 			6, 14, 46
										EMReadScreen hc_prog_elig_test_withdrawn(hc_prog_count), 		6, 15, 46
									ElseIf hc_prog_elig_major_program(hc_prog_count) = "EMA" Then
										EMReadScreen hc_prog_elig_test_parent(hc_prog_count), 		6, 11, 46
										EMReadScreen hc_prog_elig_test_residence(hc_prog_count), 	6, 12, 46
										EMReadScreen hc_prog_elig_test_verif(hc_prog_count), 		6, 13, 46
										EMReadScreen hc_prog_elig_test_withdrawn(hc_prog_count), 	6, 14, 46
									End If
								End If

								If hc_prog_elig_major_program(hc_prog_count) = "IMD" Then
									EMReadScreen hc_prog_elig_test_absence(hc_prog_count), 			6, 7, 5
									EMReadScreen hc_prog_elig_test_assets(hc_prog_count), 			6, 8, 5
									EMReadScreen hc_prog_elig_test_assistance_unit(hc_prog_count), 	6, 9, 5
									EMReadScreen hc_prog_elig_test_citizenship(hc_prog_count), 		6, 10, 5
									EMReadScreen hc_prog_elig_test_coop(hc_prog_count), 			6, 11, 5
									EMReadScreen hc_prog_elig_test_death(hc_prog_count), 			6, 12, 5
									EMReadScreen hc_prog_elig_test_fail_file(hc_prog_count), 		6, 13, 5
									EMReadScreen hc_prog_elig_test_IMD(hc_prog_count), 				6, 14, 5
									EMReadScreen hc_prog_elig_test_income(hc_prog_count), 			6, 15, 5

									EMReadScreen hc_prog_elig_test_medicare_elig(hc_prog_count), 			6, 7, 44
									EMReadScreen hc_prog_elig_test_MNSure_system(hc_prog_count), 			6, 8, 44
									EMReadScreen hc_prog_elig_test_Obligation_one_mo(hc_prog_count),		6, 9, 44
									EMReadScreen hc_prog_elig_test_obligation_six_mo(hc_prog_count), 		6, 10, 44
									EMReadScreen hc_prog_elig_test_parent(hc_prog_count), 					6, 11, 44
									EMReadScreen hc_prog_elig_test_residence(hc_prog_count), 				6, 12, 44
									EMReadScreen hc_prog_elig_test_uncompensated_transfer(hc_prog_count), 	6, 13, 44
									EMReadScreen hc_prog_elig_test_verif(hc_prog_count), 					6, 14, 44
									EMReadScreen hc_prog_elig_test_withdrawn(hc_prog_count), 				6, 15, 44
								End If

								Call write_value_and_transmit("X", 7, 3)				'Assets'
								' MsgBox "MOVING - 3"
								transmit
								' MsgBox "MOVING - 4"

								Call write_value_and_transmit("X", 10, 3)				'Cooperration'
								' MsgBox "MOVING - 5"
								Call write_value_and_transmit("X", 10, 26)				'Cooperration'
								' MsgBox "MOVING - 6"
								EMReadScreen hc_prog_elig_test_coop_pben_cash(hc_prog_count), 			6, 10, 31
								EMReadScreen hc_prog_elig_test_coop_pben_smrt(hc_prog_count), 			6, 11, 31
								transmit
								' MsgBox "MOVING - 7"
								EMReadScreen hc_prog_elig_test_coop_pben(hc_prog_count), 					6, 10, 28
								EMReadScreen hc_prog_elig_test_coop_fail_provide_info(hc_prog_count), 	6, 11, 28
								EMReadScreen hc_prog_elig_test_coop_IEVS(hc_prog_count), 					6, 12, 28
								EMReadScreen hc_prog_elig_test_coop_medical_support(hc_prog_count), 		6, 13, 28
								EMReadScreen hc_prog_elig_test_coop_other_health_ins(hc_prog_count), 		6, 14, 28
								EMReadScreen hc_prog_elig_test_coop_SSN(hc_prog_count), 					6, 15, 28
								EMReadScreen hc_prog_elig_test_coop_third_party_liability(hc_prog_count), 6, 16, 28
								transmit
								' MsgBox "MOVING - 8"

								Call write_value_and_transmit("X", 14, 3)				'Fail to File'
								' MsgBox "MOVING - 9"
								EMReadScreen hc_prog_elig_test_fail_file_HRF(hc_prog_count), 				6, 14, 33
								EMReadScreen hc_prog_elig_test_fail_file_IR(hc_prog_count), 				6, 15, 33
								EMReadScreen hc_prog_elig_test_fail_file_AR(hc_prog_count), 				6, 16, 33
								EMReadScreen hc_prog_elig_test_fail_file_ER(hc_prog_count), 				6, 17, 33
								EMReadScreen hc_prog_elig_test_fail_file_quarterly_TYMA(hc_prog_count), 	6, 18, 33
								transmit
								' MsgBox "MOVING - 10"

								If hc_prog_elig_major_program(hc_prog_count) = "MA" Then Call write_value_and_transmit("X", 14, 44)				'Verification'
								If hc_prog_elig_major_program(hc_prog_count) = "EMA" Then Call write_value_and_transmit("X", 13, 44)				'Verification'
								' MsgBox "MOVING - 11"
								EMReadScreen hc_prog_elig_test_verif_ACCT(hc_prog_count), 		6, 5, 10
								EMReadScreen hc_prog_elig_test_verif_BUSI(hc_prog_count), 		6, 6, 10
								EMReadScreen hc_prog_elig_test_verif_JOBS(hc_prog_count), 		6, 7, 10
								EMReadScreen hc_prog_elig_test_verif_IMIG_status(hc_prog_count), 	6, 8, 10
								EMReadScreen hc_prog_elig_test_verif_LUMP(hc_prog_count), 		6, 9, 10
								EMReadScreen hc_prog_elig_test_verif_OTHR(hc_prog_count), 		6, 10, 10
								EMReadScreen hc_prog_elig_test_verif_PBEN(hc_prog_count), 		6, 11, 10
								EMReadScreen hc_prog_elig_test_verif_PREG(hc_prog_count), 		6, 12, 10
								EMReadScreen hc_prog_elig_test_verif_RBIC(hc_prog_count), 		6, 13, 10
								EMReadScreen hc_prog_elig_test_verif_REST(hc_prog_count), 		6, 14, 10
								EMReadScreen hc_prog_elig_test_verif_SECU(hc_prog_count), 		6, 15, 10
								EMReadScreen hc_prog_elig_test_verif_SPON(hc_prog_count), 		6, 16, 10
								EMReadScreen hc_prog_elig_test_verif_TRAN(hc_prog_count), 		6, 17, 10
								EMReadScreen hc_prog_elig_test_verif_UNEA(hc_prog_count), 		6, 18, 10
								EMReadScreen hc_prog_elig_test_verif_cit_id(hc_prog_count), 		6, 19, 10
								EMReadScreen hc_prog_elig_test_verif_CARS(hc_prog_count), 		6, 20, 10
								transmit
								' MsgBox "MOVING - 12"

								Call write_value_and_transmit("X", 18, 3)				'Uncompensated Transfer
								' MsgBox "MOVING - 13"
								transmit
								' MsgBox "MOVING - 14"

								' Call write_value_and_transmit("X", 9, 44)				'Obligation - One Month - we don't need this
								' transmit

								transmit
								' MsgBox "MOVING - 15"

								Call write_value_and_transmit("X", 8, hc_col+4)			'Household Count'
								' MsgBox "MOVING - 16"
								EMReadScreen hc_prog_elig_hh_size(hc_prog_count), 2, 5, 68
								hh_row = 12
								Do
									EMReadScreen inc_count_ind, 1, hh_row, 61
									If inc_count_ind = "Y" Then
										EMReadScreen memb_numb_income_count, 2, hh_row, 13
										hc_prog_elig_members_whose_income_counts(hc_prog_count) = hc_prog_elig_members_whose_income_counts(hc_prog_count) & " " & memb_numb_income_count
									End If
									hh_row = hh_row + 1
									EMReadScreen next_inc_count_ind, 1, hh_row, 61
								Loop until next_inc_count_ind = " "
								hc_prog_elig_members_whose_income_counts(hc_prog_count) = trim(hc_prog_elig_members_whose_income_counts(hc_prog_count))
								hc_prog_elig_members_whose_income_counts_list(hc_prog_count) = replace(hc_prog_elig_members_whose_income_counts(hc_prog_count), " ", ",")
								hc_prog_elig_members_whose_income_counts(hc_prog_count) = split(hc_prog_elig_members_whose_income_counts(hc_prog_count), " ")
								transmit
								' MsgBox "MOVING - 17"

								If hc_prog_elig_method(hc_prog_count) <> "X" Then
									Call write_value_and_transmit("X", 9, hc_col+4)		'Budget'
									' MsgBox "MOVING - 18"
									EMReadScreen hc_prog_elig_PTMA(hc_prog_count), 1, 5, 63
									EMReadScreen hc_prog_elig_elig_standard_percent(hc_prog_count), 3, 6, 66
									EMReadScreen hc_prog_elig_basis(hc_prog_count), 20, 6, 24

									EMReadScreen budg_panel, 70, 3, 2
									' SBUD
									' CBUD
									' BBUD
									' LBUD
									' ABUD
									budget_found = False

									If InStr(budg_panel, "ABUD") <> 0 Then
										' MsgBox "ABUD"
										budget_found = true
										EMReadScreen hc_prog_elig_budg_gross_unearned(hc_prog_count), 		10, 9, 31
										EMReadScreen hc_prog_elig_budg_excluded_unearned(hc_prog_count), 	10, 10, 31
										EMReadScreen hc_prog_elig_budg_unearned_deduction(hc_prog_count), 	10, 11, 31
										EMReadScreen hc_prog_elig_budg_net_unearned_income(hc_prog_count), 	10, 12, 31

										EMReadScreen hc_prog_elig_budg_gross_earned(hc_prog_count), 		10, 14, 31
										EMReadScreen hc_prog_elig_budg_excluded_earned(hc_prog_count), 		10, 15, 31
										EMReadScreen hc_prog_elig_budg_work_exp_deduction(hc_prog_count), 	10, 16, 31
										EMReadScreen hc_prog_elig_budg_earned_disregarrd(hc_prog_count), 	10, 17, 31
										EMReadScreen hc_prog_elig_budg_dependent_care(hc_prog_count), 		10, 18, 31

										EMReadScreen hc_prog_elig_budg_earned_deduction(hc_prog_count), 	10, 9, 71
										EMReadScreen hc_prog_elig_budg_net_earned_income(hc_prog_count), 	10, 10, 71

										EMReadScreen hc_prog_elig_budg_child_sup_deduction(hc_prog_count),	10, 12, 71
										EMReadScreen hc_prog_elig_budg_deemed_income(hc_prog_count), 		10, 13, 71
										EMReadScreen hc_prog_elig_budg_total_net_income(hc_prog_count), 	10, 14, 71
										EMReadScreen hc_prog_elig_budg_income_standard(hc_prog_count), 		10, 15, 71
										EMReadScreen hc_prog_elig_budg_spenddown(hc_prog_count), 			10, 16, 71
										EMReadScreen hc_prog_elig_budg_transfer_penalty(hc_prog_count), 	10, 17, 71
										EMReadScreen hc_prog_elig_budg_total_liability(hc_prog_count), 		10, 18, 71
									End If

									If InStr(budg_panel, "BBUD") <> 0 Then
									' If hc_prog_elig_method(hc_prog_count) = "B" Then
										' MsgBox "BBUD"
										budget_found = true
										EMReadScreen hc_prog_elig_budg_gross_unearned(hc_prog_count), 		10, 8, 31
										EMReadScreen hc_prog_elig_budg_deemed_unearned(hc_prog_count), 		10, 9, 31
										EMReadScreen hc_prog_elig_budg_excluded_unearned(hc_prog_count), 	10, 10, 31
										EMReadScreen hc_prog_elig_budg_unearned_deduction(hc_prog_count), 	10, 11, 31
										EMReadScreen hc_prog_elig_budg_net_unearned_income(hc_prog_count), 	10, 12, 31

										EMReadScreen hc_prog_elig_budg_gross_earned(hc_prog_count), 		10, 8, 71
										EMReadScreen hc_prog_elig_budg_deemed_earned(hc_prog_count), 		10, 9, 71
										EMReadScreen hc_prog_elig_budg_excluded_earned(hc_prog_count), 		10, 10, 71
										EMReadScreen hc_prog_elig_budg_earned_deduction(hc_prog_count), 	10, 11, 71
										EMReadScreen hc_prog_elig_budg_net_earned_income(hc_prog_count), 	10, 12, 71

										EMReadScreen hc_prog_elig_budg_total_net_income(hc_prog_count), 	10, 14, 71
										EMReadScreen hc_prog_elig_budg_income_standard(hc_prog_count), 		10, 15, 71
										EMReadScreen hc_prog_elig_budg_spenddown(hc_prog_count), 			10, 16, 71
										EMReadScreen hc_prog_elig_budg_transfer_penalty(hc_prog_count), 	10, 17, 71
										EMReadScreen hc_prog_elig_budg_total_liability(hc_prog_count), 		10, 18, 71
									End If

									If InStr(budg_panel, "CBUD") <> 0 Then
										' MsgBox "CBUD"
										budget_found = true
										EMReadScreen hc_prog_elig_budg_gross_unearned(hc_prog_count), 		10, 8, 31
										EMReadScreen hc_prog_elig_budg_deemed_unearned(hc_prog_count), 		10, 9, 31
										EMReadScreen hc_prog_elig_budg_excluded_unearned(hc_prog_count), 	10, 10, 31
										EMReadScreen hc_prog_elig_budg_net_unearned_income(hc_prog_count), 	10, 11, 31

										EMReadScreen hc_prog_elig_budg_gross_earned(hc_prog_count), 		10, 8, 71
										EMReadScreen hc_prog_elig_budg_excluded_earned(hc_prog_count), 		10, 9, 71
										EMReadScreen hc_prog_elig_budg_net_earned_income(hc_prog_count), 	10, 10, 71

										EMReadScreen hc_prog_elig_budg_deemed_earned(hc_prog_count), 		10, 13, 71
										EMReadScreen hc_prog_elig_budg_total_net_income(hc_prog_count), 	10, 14, 71
										EMReadScreen hc_prog_elig_budg_income_standard(hc_prog_count), 		10, 15, 71
										EMReadScreen hc_prog_elig_budg_excess_income(hc_prog_count), 		10, 16, 71
									End If

									If InStr(budg_panel, "LBUD") <> 0 Then
										' MsgBox "LBUD"
										budget_found = true
										EMReadScreen hc_prog_elig_budg_gross_unearned(hc_prog_count), 			10, 8, 32
										EMReadScreen hc_prog_elig_budg_countable_earrned_income(hc_prog_count),	10, 9, 32
										EMReadScreen hc_prog_elig_budg_subtotal_countable_income(hc_prog_count),10, 10, 32
										EMReadScreen hc_prog_elig_budg_va_aid_attendincome(hc_prog_count), 		10, 11, 32
										EMReadScreen hc_prog_elig_budg_total_countable_income(hc_prog_count), 	10, 12, 32

										EMReadScreen hc_prog_elig_budg_LTC_exclusions(hc_prog_count), 			10, 14, 32
										EMReadScreen hc_prog_elig_budg_medicare_premium(hc_prog_count), 		10, 15, 32
										EMReadScreen hc_prog_elig_budg_person_clothing_needs(hc_prog_count), 	10, 16, 32
										EMReadScreen hc_prog_elig_budg_maint_needs_allowance(hc_prog_count), 	10, 17, 32
										EMReadScreen hc_prog_elig_budg_guardian_rep_payee_fee(hc_prog_count), 	10, 18, 32

										EMReadScreen hc_prog_elig_budg_spousal_allocation(hc_prog_count), 		10, 8, 70
										EMReadScreen hc_prog_elig_budg_family_allocation(hc_prog_count), 		10, 9, 70
										EMReadScreen hc_prog_elig_budg_health_ins_premium(hc_prog_count), 		10, 10, 70
										EMReadScreen hc_prog_elig_budg_other_medical_expense(hc_prog_count), 	10, 11, 70
										EMReadScreen hc_prog_elig_budg_SSI_1611_benefit(hc_prog_count), 		10, 12, 70
										EMReadScreen hc_prog_elig_budg_other_deductions(hc_prog_count), 		10, 13, 70
										EMReadScreen hc_prog_elig_budg_total_net_income(hc_prog_count), 		10, 14, 70
										EMReadScreen hc_prog_elig_budg_LTC_spenddown(hc_prog_count), 			10, 15, 70
										EMReadScreen hc_prog_elig_budg_transfer_penalty(hc_prog_count), 		10, 16, 70
										EMReadScreen hc_prog_elig_budg_total_liability(hc_prog_count), 			10, 17, 70
										EMReadScreen hc_prog_elig_budg_medical_spenddown(hc_prog_count), 		10, 18, 70
									End If

									If InStr(budg_panel, "SBUD") <> 0 Then
										' MsgBox "SBUD"
										budget_found = true
										EMReadScreen hc_prog_elig_budg_gross_unearned(hc_prog_count), 			10, 9, 31
										EMReadScreen hc_prog_elig_budg_countable_earrned_income(hc_prog_count),	10, 10, 31
										EMReadScreen hc_prog_elig_budg_subtotal_countable_income(hc_prog_count),10, 11, 31
										EMReadScreen hc_prog_elig_budg_va_aid_attendincome(hc_prog_count), 		10, 12, 31
										EMReadScreen hc_prog_elig_budg_total_countable_income(hc_prog_count), 	10, 13, 31

										EMReadScreen hc_prog_elig_budg_LTC_exclusions(hc_prog_count), 			10, 15, 31
										EMReadScreen hc_prog_elig_budg_medicare_premium(hc_prog_count), 		10, 16, 31
										EMReadScreen hc_prog_elig_budg_maint_needs_allowance(hc_prog_count), 	10, 17, 31
										EMReadScreen hc_prog_elig_budg_guardian_rep_payee_fee(hc_prog_count), 	10, 18, 31

										EMReadScreen hc_prog_elig_budg_spousal_allocation(hc_prog_count), 		10, 9, 71
										EMReadScreen hc_prog_elig_budg_family_allocation(hc_prog_count), 		10, 10, 71
										EMReadScreen hc_prog_elig_budg_health_ins_premium(hc_prog_count), 		10, 11, 71
										EMReadScreen hc_prog_elig_budg_other_medical_expense(hc_prog_count), 	10, 12, 71
										EMReadScreen hc_prog_elig_budg_SSI_1611_benefit(hc_prog_count), 		10, 13, 71
										EMReadScreen hc_prog_elig_budg_other_deductions(hc_prog_count), 		10, 14, 71
										EMReadScreen hc_prog_elig_budg_total_net_income(hc_prog_count), 		10, 15, 71
										EMReadScreen hc_prog_elig_budg_waiver_obligation(hc_prog_count),	 	10, 16, 71
										EMReadScreen hc_prog_elig_budg_transfer_penalty(hc_prog_count), 		10, 17, 71
										EMReadScreen hc_prog_elig_budg_total_liability(hc_prog_count), 			10, 18, 71
									End If
									If budget_found = false Then MsgBox "Budget not coded:" & vbCr & budg_panel

									transmit
									' MsgBox "MOVING - 19"
								End If

								Call write_value_and_transmit("X", 18, 3)				'MOBL
								' MsgBox "MOVING - 20"
								EMReadScreen hc_prog_elig_mobl_result(hc_prog_count), 4, 6, 32
								EMReadScreen hc_prog_elig_mobl_type(hc_prog_count), 	18, 6, 39
								EMReadScreen hc_prog_elig_mobl_period(hc_prog_count), 13, 6, 61
								mobl_row = 6
								Do
									EMReadScreen mobl_ref_numb, 2, mobl_row, 6
									If mobl_ref_numb = hc_elig_ref_numbs(hc_prog_count) Then Exit Do
									mobl_row = mobl_row + 1
								Loop until mobl_ref_numb = "  "
								Call write_value_and_transmit("X", mobl_row, 3)				'MOBL
								Do
									' MsgBox "MOVING - 21"
									EMReadScreen spenddown_header, 75, 3, 2
									spenddown_header = trim(spenddown_header)
									If spenddown_header = "Community Spenddown Results (SPDN)" Then
										EMReadScreen hc_prog_elig_spdn_option(hc_prog_count), 	2, 4, 59
										EMReadScreen hc_prog_elig_spdn_type(hc_prog_count), 		1, 5, 14
										EMReadScreen hc_prog_elig_spdn_method(hc_prog_count), 	1, 5, 45
										EMReadScreen hc_prog_elig_spdn_covered_pop(hc_prog_count), 1, 5, 68

										mobl_col = 21
										Do
											EMReadScreen mobl_mo, 2, 7, mobl_col
											EMReadScreen mobl_yr, 2, 7, mobl_col
											If budg_mo = elig_footer_month AND budg_yr = elig_footer_year Then
												EMReadScreen hc_prog_elig_original_monthly_spdn(hc_prog_count), 			10, 8, mobl_col-5
												EMReadScreen hc_prog_elig_monthly_spdn_counted_bills(hc_prog_count), 		10, 9, mobl_col-5
												EMReadScreen hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count),	5, 10, mobl_col
												EMReadScreen hc_prog_elig_monthly_spdn_recipient_amount(hc_prog_count), 	10, 11, mobl_col-5
												EMReadScreen hc_prog_elig_monthly_spdn_balance(hc_prog_count), 				10, 12, mobl_col-5

												If hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count) <> "__ __" Then
													hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count) = replace(hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count), " ", "/")
													hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count) = hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count) & "/" & elig_footer_year
												Else
													hc_prog_elig_monthly_spdn_satisfaction_date(hc_prog_count) = ""
												End If
											End If
											mobl_col = mobl_col + 11
										Loop until mobl_col = 87
										EMReadScreen hc_prog_elig_oiginal_six_mo_spdn(hc_prog_count), 			10, 15, 45
										EMReadScreen hc_prog_elig_six_mo_spdn_counted_bills(hc_prog_count), 		10, 16, 45
										EMReadScreen hc_prog_elig_six_mo_spnd_satisfaction_date(hc_prog_count), 	8, 17, 45
										EMReadScreen hc_prog_elig_six_mo_spdn_recipient_amount(hc_prog_count), 	10, 18, 45
										EMReadScreen hc_prog_elig_six_mo_spdn_unused_balance(hc_prog_count), 		10, 19, 45
									ElseIf spenddown_header = "SIS-EW Waiver Obligation Results (EWWO)" Then
										'2506494
									ElseIf spenddown_header = "Long Term Care/Medical Spenddown Results (LTCS)" Then
										'804476
									Else
										MsgBox spenddown_header
									End If
									transmit
									EMReadScreen back_to_MOBL_check, 4,	 3, 49
								Loop until back_to_MOBL_check = "MOBL"
								' MsgBox "MOVING - 22"
								PF3

								Call write_value_and_transmit("X", 18, 34)				'Cert Period Amount'
								' MsgBox "MOVING - 22.5"
								EMReadScreen cert_pd_pop_up_check, 27, 5, 13
								If cert_pd_pop_up_check = "Certification Period Amount" Then
									EMReadScreen hc_prog_elig_cert_prd_total_net_income(hc_prog_count), 	10, 7, 34
									EMReadScreen hc_prog_elig_cert_prd_total_standard(hc_prog_count), 		10, 8, 34
									EMReadScreen hc_prog_elig_cert_prd_total_excess_income(hc_prog_count), 	10, 9, 34
									transmit
								End If
								EMWriteScreen " ", 18, 34
								' MsgBox "MOVING - 23"

								transmit
								' MsgBox "MOVING - 24"

								EMReadScreen hc_prog_elig_app_date(hc_prog_count), 8, 4, 73

								EMReadScreen hc_prog_elig_curr_prog_status(hc_prog_count), 10, 8, 34
								EMReadScreen hc_prog_elig_elig_result(hc_prog_count), 	10, 9, 34
								EMReadScreen hc_prog_elig_elig_begin_date(hc_prog_count), 8, 10, 34
								EMReadScreen hc_prog_elig_HRF_reporting(hc_prog_count), 	10, 11, 34
								EMReadScreen hc_prog_elig_ER_date(hc_prog_count), 		8, 12, 34
								EMReadScreen hc_prog_elig_SR_date(hc_prog_count), 		8, 13, 34
								If trim(hc_prog_elig_SR_date(hc_prog_count)) <> "" Then EMReadScreen hc_prog_elig_SR_date(hc_prog_count), 		8, 14, 34
								EMReadScreen hc_prog_elig_TYMA_begin_date(hc_prog_count), 8, 15, 34

								EMReadScreen hc_prog_elig_responsible_county(hc_prog_count), 	2, 8, 76
								EMReadScreen hc_prog_elig_servicing_county(hc_prog_count), 	2, 9, 76
								EMReadScreen hc_prog_elig_source_of_info(hc_prog_count), 		4, 10, 76

								EMReadScreen hc_prog_elig_TYMA_type(hc_prog_count), 2, 15, 76

								EMReadScreen hc_prog_elig_worker_msg_one(hc_prog_count), 78, 18, 3
								EMReadScreen hc_prog_elig_worker_msg_two(hc_prog_count), 78, 19, 3

								transmit
								' MsgBox "MOVING - 25"

								Exit Do
							End If
							hc_col = hc_col + 11
							If hc_col = 85 Then hc_prog_elig_appd(hc_prog_count) = False
						Loop until hc_col = 83
					End If

					If hc_prog_elig_major_program(hc_prog_count) = "QMB" or hc_prog_elig_major_program(hc_prog_count) = "SLMB" or hc_prog_elig_major_program(hc_prog_count) = "QI1" Then
						' MsgBox hc_prog_elig_major_program(hc_prog_count)
						' EmReadScreen hc_elig_membs_prog_one
						EMReadScreen hc_prog_elig_elig_type(hc_prog_count), 		2, 6, 56
						EMReadScreen hc_prog_elig_elig_standard(hc_prog_count), 	1, 6, 64
						EMReadScreen hc_prog_elig_elig_standard_percent(hc_prog_count), 3, 6, 66
						EMReadScreen hc_prog_elig_basis(hc_prog_count), 			15, 6, 27

						EMReadScreen hc_prog_elig_budg_gross_unearned(hc_prog_count), 		10, 9, 31
						EMReadScreen hc_prog_elig_budg_deemed_unearned(hc_prog_count), 		10, 10, 31
						EMReadScreen hc_prog_elig_budg_excluded_unearned(hc_prog_count), 		10, 11, 31
						EMReadScreen hc_prog_elig_budg_unearned_deduction(hc_prog_count), 	10, 12, 31
						EMReadScreen hc_prog_elig_budg_standard_disregard(hc_prog_count), 	10, 13, 31
						EMReadScreen hc_prog_elig_budg_net_unearned_income(hc_prog_count), 	10, 14, 31

						EMReadScreen hc_prog_elig_budg_gross_earned(hc_prog_count), 		10, 9, 71
						EMReadScreen hc_prog_elig_budg_deemed_earned(hc_prog_count), 		10, 10, 71
						EMReadScreen hc_prog_elig_budg_excluded_earned(hc_prog_count), 	10, 11, 71
						EMReadScreen hc_prog_elig_budg_earned_deduction(hc_prog_count), 	10, 12, 71
						EMReadScreen hc_prog_elig_budg_net_earned_income(hc_prog_count), 	10, 13, 71

						EMReadScreen hc_prog_elig_budg_total_net_income(hc_prog_count), 	10, 15, 71
						EMReadScreen hc_prog_elig_budg_income_standard(hc_prog_count), 	10, 16, 71
						EMReadScreen hc_prog_elig_budg_excess_income(hc_prog_count), 		10, 17, 71

						Call write_value_and_transmit("X", 5, 66)			'Household Count'
						' MsgBox "MOVING - 26"
						EMReadScreen hc_prog_elig_hh_size(hc_prog_count), 2, 5, 68
						hh_row = 12
						Do
							EMReadScreen inc_count_ind, 1, hh_row, 61
							If inc_count_ind = "Y" Then
								EMReadScreen memb_numb_income_count, 2, hh_row, 13
								hc_prog_elig_members_whose_income_counts(hc_prog_count) = hc_prog_elig_members_whose_income_counts(hc_prog_count) & " " & memb_numb_income_count
							End If
							hh_row = hh_row + 1
							EMReadScreen next_inc_count_ind, 1, hh_row, 61
						Loop until next_inc_count_ind = " "
						hc_prog_elig_members_whose_income_counts(hc_prog_count) = trim(hc_prog_elig_members_whose_income_counts(hc_prog_count))
						hc_prog_elig_members_whose_income_counts_list(hc_prog_count) = replace(hc_prog_elig_members_whose_income_counts(hc_prog_count), " ", ",")
						hc_prog_elig_members_whose_income_counts(hc_prog_count) = split(hc_prog_elig_members_whose_income_counts(hc_prog_count), " ")
						transmit
						' MsgBox "MOVING - 27"

						transmit
						' MsgBox "MOVING - 28"

						EMReadScreen hc_prog_elig_test_absence(hc_prog_count), 				6, 6, 5
						EMReadScreen hc_prog_elig_test_after_processing_month(hc_prog_count),	6, 7, 5
						EMReadScreen hc_prog_elig_test_assets(hc_prog_count), 				6, 8, 5
						EMReadScreen hc_prog_elig_test_assistance_unit(hc_prog_count), 		6, 9, 5
						EMReadScreen hc_prog_elig_test_basis_for_other_prog(hc_prog_count), 	6, 10, 5
						EMReadScreen hc_prog_elig_test_citizenship(hc_prog_count), 			6, 11, 5
						EMReadScreen hc_prog_elig_test_coop(hc_prog_count), 					6, 12, 5
						EMReadScreen hc_prog_elig_test_correctional_faci(hc_prog_count), 		6, 13, 5

						EMReadScreen hc_prog_elig_test_death(hc_prog_count), 					6, 6, 46
						EMReadScreen hc_prog_elig_test_fail_file(hc_prog_count), 				6, 7, 46
						EMReadScreen hc_prog_elig_test_income(hc_prog_count), 				6, 8, 46
						EMReadScreen hc_prog_elig_test_medicare_part_a(hc_prog_count),		6, 9, 46
						EMReadScreen hc_prog_elig_test_residence(hc_prog_count), 				6, 10, 46
						EMReadScreen hc_prog_elig_test_verif(hc_prog_count), 					6, 11, 46
						EMReadScreen hc_prog_elig_test_withdrawn(hc_prog_count), 				6, 12, 46

						EMReadScreen hc_prog_elig_test_uncompensated_transfer(hc_prog_count), 6, 17, 5

						transmit
						' MsgBox "MOVING - 29"

						EMReadScreen hc_prog_elig_app_date(hc_prog_count), 8, 4, 73

						EMReadScreen hc_prog_elig_curr_prog_status(hc_prog_count), 10, 8, 34
						EMReadScreen hc_prog_elig_elig_result(hc_prog_count), 	10, 9, 34
						EMReadScreen hc_prog_elig_elig_begin_date(hc_prog_count), 8, 10, 34
						EMReadScreen hc_prog_elig_ER_date(hc_prog_count), 		8, 11, 34
						EMReadScreen hc_prog_elig_SR_date(hc_prog_count), 		8, 12, 34
						If trim(hc_prog_elig_SR_date(hc_prog_count)) <> "" Then EMReadScreen hc_prog_elig_SR_date(hc_prog_count), 		8, 13, 34
						EMReadScreen hc_prog_elig_source_of_info(hc_prog_count), 	4, 14, 34

						EMReadScreen hc_prog_elig_responsible_county(hc_prog_count), 	2, 8, 78
						EMReadScreen hc_prog_elig_servicing_county(hc_prog_count), 	2, 9, 78

						EMReadScreen hc_prog_elig_worker_msg_one(hc_prog_count), 78, 18, 3
						EMReadScreen hc_prog_elig_worker_msg_two(hc_prog_count), 78, 19, 3

						transmit
						' MsgBox "MOVING - 30"

					End If
				' Else
				' 	hc_prog_elig_appd(hc_prog_count) = False
				' End If
			End If

			' EMReadScreen next_ref_numb, 2, hc_row+1, 3
			' If next_ref_numb = "  " Then
			' 	hc_row = hc_row + 1
			'
			' 	EMReadScreen clt_hc_prog, 4, hc_row, 28
			' 	If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "    " Then
			'
			' 		EMReadScreen prog_status, 3, hc_row, 68
			' 		If prog_status <> "APP" Then                        'Finding the approved version
			' 			EMReadScreen total_versions, 2, hc_row, 64
			' 			If total_versions = "01" Then
			' 				hc_elig_membs_prog_two_appd(hc_prog_count) = False
			' 			Else
			' 				EMReadScreen current_version, 2, hc_row, 58
			' 				If current_version = "01" Then
			' 					hc_elig_membs_prog_two_appd(hc_prog_count) = False
			' 				Else
			' 					prev_version = right ("00" & abs(current_version) - 1, 2)
			' 					EMWriteScreen prev_version, hc_row, 58
			' 					transmit
			' 					hc_elig_membs_prog_two_appd(hc_prog_count) = True
			' 				End If
			'
			' 			End If
			' 		Else
			' 			hc_elig_membs_prog_two_appd(hc_prog_count) = True
			' 		End If
			' 	Else
			' 		hc_elig_membs_prog_two_appd(hc_prog_count) = False
			' 	End If
			'
			' 	If hc_elig_membs_prog_two_appd(hc_prog_count) = True Then
			' 		EMReadScreen hc_elig_membs_prog_two_process_date(hc_prog_count), 8, 2, 73
			' 		hc_elig_membs_prog_two_process_date(hc_prog_count) = DateAdd("d", 0, hc_elig_membs_prog_two_process_date(hc_prog_count))
			'
			' 		If DateDiff("'d", hc_elig_membs_prog_two_process_date(hc_prog_count), date) = 0 Then
			' 			EMReadScreen hc_elig_membs_program_two(hc_prog_count), 4, hc_row, 28
			' 			EMReadScreen hc_elig_membs_prog_two_eligibility_result(hc_prog_count), 8, hc_row, 41
			' 			EMReadScreen hc_elig_membs_prog_two_status(hc_prog_count), 8, hc_row, 50
			' 			EMReadScreen hc_elig_membs_prog_two_app_indc(hc_prog_count), 6, hc_row, 68
			' 			EMReadScreen hc_elig_membs_prog_two_magi_excempt(hc_prog_count), 6, hc_row, 74
			' 		Else
			'
			' 		End If
			' 	End If
			' End If

			hc_prog_count = hc_prog_count + 1
			hc_row = hc_row + 1
			EMReadScreen next_ref_numb, 2, hc_row, 3
			EMReadScreen next_maj_prog, 4, hc_row, 28
			' MsgBox "Row: " & hc_row & vbCr & "Next Ref Numb: " & next_ref_numb & vbCr & "Next Major Prog: " & next_maj_prog
		Loop until next_ref_numb = "  " and next_maj_prog = "    "
		Call back_to_SELF
	end sub
end class

class stat_detail
	public footer_month
	public footer_year
	public stat_prog_cash_I_appl_date
	public stat_prog_cash_I_elig_begin_date
	public stat_prog_cash_I_interview_date
	public stat_prog_cash_I_prog
	public stat_prog_cash_I_status
	public stat_prog_cash_II_appl_date
	public stat_prog_cash_II_elig_begin_date
	public stat_prog_cash_II_interview_date
	public stat_prog_cash_II_prog
	public stat_prog_cash_II_status
	public stat_prog_emer_appl_date
	public stat_prog_emer_elig_begin_date
	public stat_prog_emer_interview_date
	public stat_prog_emer_prog
	public stat_prog_emer_status
	public stat_prog_grh_appl_date
	public stat_prog_grh_elig_begin_date
	public stat_prog_grh_interview_date
	public stat_prog_grh_status
	public stat_prog_snap_appl_date
	public stat_prog_snap_elig_begin_date
	public stat_prog_snap_interview_date
	public stat_prog_snap_status
	public stat_prog_ive_appl_date
	public stat_prog_ive_elig_begin_date
	public stat_prog_ive_interview_date
	public stat_prog_ive_status
	public stat_prog_hc_appl_date
	public stat_prog_hc_elig_begin_date
	public stat_prog_hc_interview_date
	public stat_prog_hc_status


	public stat_hest_persons_paying_list
	public stat_hest_retro_heat_air_yn
	public stat_hest_retro_heat_air_fs_units
	public stat_hest_retro_electric_yn
	public stat_hest_retro_electric_fs_units
	public stat_hest_retro_phone_yn
	public stat_hest_retro_phone_fs_units
	public stat_hest_prosp_heat_air_yn
	public stat_hest_prosp_heat_air_fs_units
	public stat_hest_prosp_electric_yn
	public stat_hest_prosp_electric_fs_units
	public stat_hest_prosp_phone_yn
	public stat_hest_prosp_phone_fs_units

	public stat_revw_cash_code
	public stat_next_cash_revw_date
	public stat_next_cash_revw_process
	public stat_last_cash_revw_date
	public stat_last_cash_revw_process
	public stat_revw_snap_code
	public stat_next_snap_revw_date
	public stat_next_snap_revw_process
	public stat_last_snap_revw_date
	public stat_last_snap_revw_process
	public stat_revw_hc_code
	public stat_next_hc_revw_date
	public stat_next_hc_revw_process
	public stat_last_hc_revw_date
	public stat_last_hc_revw_process
	public stat_revw_form_recvd_date
	public stat_revw_interview_date
	public stat_mont_cash_status
	public stat_mont_snap_status
	public stat_mont_hc_status
	public stat_mont_form_recvd_date
	public stat_shel_prosp_all_total
	public stat_hest_retro_heat_air_amount
	public stat_hest_retro_electric_amount
	public stat_hest_retro_phone_amount
	public stat_hest_prosp_heat_air_amount
	public stat_hest_prosp_electric_amount
	public stat_hest_prosp_phone_amount
	public stat_hest_retro_all
	public stat_hest_prosp_all
	public stat_hest_retro_list
	public stat_hest_prosp_list

	public stat_memb_ref_numb()
	public stat_memb_first_name()
	public stat_memb_last_name()
	public stat_memb_middle_initial()
	public stat_memb_full_name()
	public stat_memb_full_name_no_initial()
	public stat_memb_full_name_last_name_first()
	public stat_memb_full_name_last_name_first_no_mi()
	public stat_memb_age()
	public stat_memb_id_verif_code()
	public stat_memb_id_verif_info()
	public stat_memb_rel_to_applct_code()
	public stat_memb_rel_to_applct_info()
	public stat_memi_spouse_ref_numb()
	public stat_memi_citizenship_yn()
	public stat_memi_citizenship_verif_code()
	public stat_memi_citizenship_verif_info()
	public stat_jobs_one_exists()
	public stat_jobs_one_job_ended()
	public stat_jobs_one_job_counted()
	public stat_jobs_one_inc_type()
	public stat_jobs_one_sub_inc_type()
	public stat_jobs_one_verif_code()
	public stat_jobs_one_verif_info()
	public stat_jobs_one_employer_name()
	public stat_jobs_one_inc_start_date()
	public stat_jobs_one_inc_end_date()
	public stat_jobs_one_main_pay_freq()
	public stat_jobs_one_snap_pic_pay_freq()
	public stat_jobs_one_snap_pic_ave_hrs_per_pay()
	public stat_jobs_one_snap_pic_ave_inc_per_pay()
	public stat_jobs_one_snap_pic_prosp_monthly_inc()
	public stat_jobs_one_grh_pic_pay_freq()
	public stat_jobs_one_grh_pic_ave_inc_per_pay()
	public stat_jobs_one_grh_pic_prosp_monthly_inc()
	public stat_jobs_two_exists()
	public stat_jobs_two_job_ended()
	public stat_jobs_two_job_counted()
	public stat_jobs_two_inc_type()
	public stat_jobs_two_sub_inc_type()
	public stat_jobs_two_verif_code()
	public stat_jobs_two_verif_info()
	public stat_jobs_two_employer_name()
	public stat_jobs_two_inc_start_date()
	public stat_jobs_two_inc_end_date()
	public stat_jobs_two_main_pay_freq()
	public stat_jobs_two_snap_pic_pay_freq()
	public stat_jobs_two_snap_pic_ave_hrs_per_pay()
	public stat_jobs_two_snap_pic_ave_inc_per_pay()
	public stat_jobs_two_snap_pic_prosp_monthly_inc()
	public stat_jobs_two_grh_pic_pay_freq()
	public stat_jobs_two_grh_pic_ave_inc_per_pay()
	public stat_jobs_two_grh_pic_prosp_monthly_inc()
	public stat_jobs_three_exists()
	public stat_jobs_three_job_ended()
	public stat_jobs_three_job_counted()
	public stat_jobs_three_inc_type()
	public stat_jobs_three_sub_inc_type()
	public stat_jobs_three_verif_code()
	public stat_jobs_three_verif_info()
	public stat_jobs_three_employer_name()
	public stat_jobs_three_inc_start_date()
	public stat_jobs_three_inc_end_date()
	public stat_jobs_three_main_pay_freq()
	public stat_jobs_three_snap_pic_pay_freq()
	public stat_jobs_three_snap_pic_ave_hrs_per_pay()
	public stat_jobs_three_snap_pic_ave_inc_per_pay()
	public stat_jobs_three_snap_pic_prosp_monthly_inc()
	public stat_jobs_three_grh_pic_pay_freq()
	public stat_jobs_three_grh_pic_ave_inc_per_pay()
	public stat_jobs_three_grh_pic_prosp_monthly_inc()
	public stat_jobs_four_exists()
	public stat_jobs_four_job_ended()
	public stat_jobs_four_job_counted()
	public stat_jobs_four_inc_type()
	public stat_jobs_four_sub_inc_type()
	public stat_jobs_four_verif_code()
	public stat_jobs_four_verif_info()
	public stat_jobs_four_employer_name()
	public stat_jobs_four_inc_start_date()
	public stat_jobs_four_inc_end_date()
	public stat_jobs_four_main_pay_freq()
	public stat_jobs_four_snap_pic_pay_freq()
	public stat_jobs_four_snap_pic_ave_hrs_per_pay()
	public stat_jobs_four_snap_pic_ave_inc_per_pay()
	public stat_jobs_four_snap_pic_prosp_monthly_inc()
	public stat_jobs_four_grh_pic_pay_freq()
	public stat_jobs_four_grh_pic_ave_inc_per_pay()
	public stat_jobs_four_grh_pic_prosp_monthly_inc()
	public stat_jobs_five_exists()
	public stat_jobs_five_job_ended()
	public stat_jobs_five_job_counted()
	public stat_jobs_five_inc_type()
	public stat_jobs_five_sub_inc_type()
	public stat_jobs_five_verif_code()
	public stat_jobs_five_verif_info()
	public stat_jobs_five_employer_name()
	public stat_jobs_five_inc_start_date()
	public stat_jobs_five_inc_end_date()
	public stat_jobs_five_main_pay_freq()
	public stat_jobs_five_snap_pic_pay_freq()
	public stat_jobs_five_snap_pic_ave_hrs_per_pay()
	public stat_jobs_five_snap_pic_ave_inc_per_pay()
	public stat_jobs_five_snap_pic_prosp_monthly_inc()
	public stat_jobs_five_grh_pic_pay_freq()
	public stat_jobs_five_grh_pic_ave_inc_per_pay()
	public stat_jobs_five_grh_pic_prosp_monthly_inc()
	public stat_busi_one_exists()
	public stat_busi_one_type()
	public stat_busi_one_counted()
	public stat_busi_one_type_info()
	public stat_busi_one_inc_start_date()
	public stat_busi_one_inc_end_date()
	public stat_busi_one_method()
	public stat_busi_one_method_date()
	public stat_busi_one_snap_retro_net_inc()
	public stat_busi_one_snap_prosp_net_inc()
	public stat_busi_one_snap_retro_gross_inc()
	public stat_busi_one_snap_retro_expenses()
	public stat_busi_one_snap_income_verif_code()
	public stat_busi_one_snap_income_verif_info()
	public stat_busi_one_snap_prosp_gross_inc()
	public stat_busi_one_snap_prosp_expenses()
	public stat_busi_one_snap_expense_verif_code()
	public stat_busi_one_snap_expense_verif_info()
	public stat_busi_two_exists()
	public stat_busi_two_type()
	public stat_busi_two_counted()
	public stat_busi_two_type_info()
	public stat_busi_two_inc_start_date()
	public stat_busi_two_inc_end_date()
	public stat_busi_two_method()
	public stat_busi_two_method_date()
	public stat_busi_two_snap_retro_net_inc()
	public stat_busi_two_snap_prosp_net_inc()
	public stat_busi_two_snap_retro_gross_inc()
	public stat_busi_two_snap_retro_expenses()
	public stat_busi_two_snap_income_verif_code()
	public stat_busi_two_snap_income_verif_info()
	public stat_busi_two_snap_prosp_gross_inc()
	public stat_busi_two_snap_prosp_expenses()
	public stat_busi_two_snap_expense_verif_code()
	public stat_busi_two_snap_expense_verif_info()
	public stat_busi_three_exists()
	public stat_busi_three_type()
	public stat_busi_three_counted()
	public stat_busi_three_type_info()
	public stat_busi_three_inc_start_date()
	public stat_busi_three_inc_end_date()
	public stat_busi_three_method()
	public stat_busi_three_method_date()
	public stat_busi_three_snap_retro_net_inc()
	public stat_busi_three_snap_prosp_net_inc()
	public stat_busi_three_snap_retro_gross_inc()
	public stat_busi_three_snap_retro_expenses()
	public stat_busi_three_snap_income_verif_code()
	public stat_busi_three_snap_income_verif_info()
	public stat_busi_three_snap_prosp_gross_inc()
	public stat_busi_three_snap_prosp_expenses()
	public stat_busi_three_snap_expense_verif_code()
	public stat_busi_three_snap_expense_verif_info()
	public stat_unea_one_exists()
	public stat_unea_one_counted()
	public stat_unea_one_type_code()
	public stat_unea_one_type_info()
	public stat_unea_one_verif_code()
	public stat_unea_one_verif_info()
	public stat_unea_one_inc_start_date()
	public stat_unea_one_inc_end_date()
	public stat_unea_one_snap_pic_pay_freq()
	public stat_unea_one_snap_pic_ave_inc_per_pay()
	public stat_unea_one_snap_pic_prosp_monthly_inc()
	public stat_unea_two_exists()
	public stat_unea_two_counted()
	public stat_unea_two_type_code()
	public stat_unea_two_type_info()
	public stat_unea_two_verif_code()
	public stat_unea_two_verif_info()
	public stat_unea_two_inc_start_date()
	public stat_unea_two_inc_end_date()
	public stat_unea_two_snap_pic_pay_freq()
	public stat_unea_two_snap_pic_ave_inc_per_pay()
	public stat_unea_two_snap_pic_prosp_monthly_inc()
	public stat_unea_three_exists()
	public stat_unea_three_counted()
	public stat_unea_three_type_code()
	public stat_unea_three_type_info()
	public stat_unea_three_verif_code()
	public stat_unea_three_verif_info()
	public stat_unea_three_inc_start_date()
	public stat_unea_three_inc_end_date()
	public stat_unea_three_snap_pic_pay_freq()
	public stat_unea_three_snap_pic_ave_inc_per_pay()
	public stat_unea_three_snap_pic_prosp_monthly_inc()
	public stat_unea_four_exists()
	public stat_unea_four_counted()
	public stat_unea_four_type_code()
	public stat_unea_four_type_info()
	public stat_unea_four_verif_code()
	public stat_unea_four_verif_info()
	public stat_unea_four_inc_start_date()
	public stat_unea_four_inc_end_date()
	public stat_unea_four_snap_pic_pay_freq()
	public stat_unea_four_snap_pic_ave_inc_per_pay()
	public stat_unea_four_snap_pic_prosp_monthly_inc()
	public stat_unea_five_exists()
	public stat_unea_five_counted()
	public stat_unea_five_type_code()
	public stat_unea_five_type_info()
	public stat_unea_five_verif_code()
	public stat_unea_five_verif_info()
	public stat_unea_five_inc_start_date()
	public stat_unea_five_inc_end_date()
	public stat_unea_five_snap_pic_pay_freq()
	public stat_unea_five_snap_pic_ave_inc_per_pay()
	public stat_unea_five_snap_pic_prosp_monthly_inc()
	public stat_acct_one_exists()
	public stat_acct_one_type()
	public stat_acct_one_balence()
	public stat_acct_one_count_snap_yn()
	public stat_acct_two_exists()
	public stat_acct_two_type()
	public stat_acct_two_balence()
	public stat_acct_two_count_snap_yn()
	public stat_acct_three_exists()
	public stat_acct_three_type()
	public stat_acct_three_balence()
	public stat_acct_three_count_snap_yn()
	public stat_acct_four_exists()
	public stat_acct_four_type()
	public stat_acct_four_balence()
	public stat_acct_four_count_snap_yn()
	public stat_acct_five_exists()
	public stat_acct_five_type()
	public stat_acct_five_balence()
	public stat_acct_five_count_snap_yn()
	public stat_shel_exists()
	public stat_shel_subsidized_yn()
	public stat_shel_shared_yn()
	public stat_shel_paid_to()
	public stat_shel_retro_rent_amount()
	public stat_shel_retro_rent_verif_code()
	public stat_shel_retro_rent_verif_info()
	public stat_shel_prosp_rent_amount()
	public stat_shel_prosp_rent_verif_code()
	public stat_shel_prosp_rent_verif_info()
	public stat_shel_retro_lot_rent_amount()
	public stat_shel_retro_lot_rent_verif_code()
	public stat_shel_retro_lot_rent_verif_info()
	public stat_shel_prosp_lot_rent_amount()
	public stat_shel_prosp_lot_rent_verif_code()
	public stat_shel_prosp_lot_rent_verif_info()
	public stat_shel_retro_mortgage_amount()
	public stat_shel_retro_mortgage_verif_code()
	public stat_shel_retro_mortgage_verif_info()
	public stat_shel_prosp_mortgage_amount()
	public stat_shel_prosp_mortgage_verif_code()
	public stat_shel_prosp_mortgage_verif_info()
	public stat_shel_retro_insurance_amount()
	public stat_shel_retro_insurance_verif_code()
	public stat_shel_retro_insurance_verif_info()
	public stat_shel_prosp_insurance_amount()
	public stat_shel_prosp_insurance_verif_code()
	public stat_shel_prosp_insurance_verif_info()
	public stat_shel_retro_taxes_amount()
	public stat_shel_retro_taxes_verif_code()
	public stat_shel_retro_taxes_verif_info()
	public stat_shel_prosp_taxes_amount()
	public stat_shel_prosp_taxes_verif_code()
	public stat_shel_prosp_taxes_verif_info()
	public stat_shel_retro_room_amount()
	public stat_shel_retro_room_verif_code()
	public stat_shel_retro_room_verif_info()
	public stat_shel_prosp_room_amount()
	public stat_shel_prosp_room_verif_code()
	public stat_shel_prosp_room_verif_info()
	public stat_shel_retro_garage_amount()
	public stat_shel_retro_garage_verif_code()
	public stat_shel_retro_garage_verif_info()
	public stat_shel_prosp_garage_amount()
	public stat_shel_prosp_garage_verif_code()
	public stat_shel_prosp_garage_verif_info()
	public stat_shel_retro_subsidy_amount()
	public stat_shel_retro_subsidy_verif_code()
	public stat_shel_retro_subsidy_verif_info()
	public stat_shel_prosp_subsidy_amount()
	public stat_shel_prosp_subsidy_verif_code()
	public stat_shel_prosp_subsidy_verif_info()

	public stat_disq_one_exists()
	public stat_disq_one_program()
	public stat_disq_one_type_code()
	public stat_disq_one_type_info()
	public stat_disq_one_begin_date()
	public stat_disq_one_end_date()
	public stat_disq_one_cure_reason_code()
	public stat_disq_one_cure_reason_info()
	public stat_disq_one_fraud_determination_date()
	public stat_disq_one_county_of_fraud()
	public stat_disq_one_state_of_fraud()
	public stat_disq_one_SNAP_trafficking_yn()
	public stat_disq_one_SNAP_offense_code()
	public stat_disq_one_SNAP_offense_info()
	public stat_disq_one_source()
	public stat_disq_one_active()

	public stat_disq_two_exists()
	public stat_disq_two_program()
	public stat_disq_two_type_code()
	public stat_disq_two_type_info()
	public stat_disq_two_begin_date()
	public stat_disq_two_end_date()
	public stat_disq_two_cure_reason_code()
	public stat_disq_two_cure_reason_info()
	public stat_disq_two_fraud_determination_date()
	public stat_disq_two_county_of_fraud()
	public stat_disq_two_state_of_fraud()
	public stat_disq_two_SNAP_trafficking_yn()
	public stat_disq_two_SNAP_offense_code()
	public stat_disq_two_SNAP_offense_info()
	public stat_disq_two_source()
	public stat_disq_two_active()

	public stat_disq_three_exists()
	public stat_disq_three_program()
	public stat_disq_three_type_code()
	public stat_disq_three_type_info()
	public stat_disq_three_begin_date()
	public stat_disq_three_end_date()
	public stat_disq_three_cure_reason_code()
	public stat_disq_three_cure_reason_info()
	public stat_disq_three_fraud_determination_date()
	public stat_disq_three_county_of_fraud()
	public stat_disq_three_state_of_fraud()
	public stat_disq_three_SNAP_trafficking_yn()
	public stat_disq_three_SNAP_offense_code()
	public stat_disq_three_SNAP_offense_info()
	public stat_disq_three_source()
	public stat_disq_three_active()

	public stat_disq_four_exists()
	public stat_disq_four_program()
	public stat_disq_four_type_code()
	public stat_disq_four_type_info()
	public stat_disq_four_begin_date()
	public stat_disq_four_end_date()
	public stat_disq_four_cure_reason_code()
	public stat_disq_four_cure_reason_info()
	public stat_disq_four_fraud_determination_date()
	public stat_disq_four_county_of_fraud()
	public stat_disq_four_state_of_fraud()
	public stat_disq_four_SNAP_trafficking_yn()
	public stat_disq_four_SNAP_offense_code()
	public stat_disq_four_SNAP_offense_info()
	public stat_disq_four_source()
	public stat_disq_four_active()

	public stat_disq_five_exists()
	public stat_disq_five_program()
	public stat_disq_five_type_code()
	public stat_disq_five_type_info()
	public stat_disq_five_begin_date()
	public stat_disq_five_end_date()
	public stat_disq_five_cure_reason_code()
	public stat_disq_five_cure_reason_info()
	public stat_disq_five_fraud_determination_date()
	public stat_disq_five_county_of_fraud()
	public stat_disq_five_state_of_fraud()
	public stat_disq_five_SNAP_trafficking_yn()
	public stat_disq_five_SNAP_offense_code()
	public stat_disq_five_SNAP_offense_info()
	public stat_disq_five_source()
	public stat_disq_five_active()


	public sub gather_stat_info()
		MAXIS_footer_month = footer_month
		MAXIS_footer_year = footer_year

		current_month = footer_month & "/1/" & footer_year
		current_month = DateAdd("d", 0, current_month)

		EMReadScreen stat_prog_cash_I_appl_date, 		8, 6, 33
		EMReadScreen stat_prog_cash_I_elig_begin_date, 	8, 6, 44
		EMReadScreen stat_prog_cash_I_interview_date, 	8, 6, 55
		EMReadScreen stat_prog_cash_I_prog, 			2, 6, 67
		EMReadScreen stat_prog_cash_I_status, 			4, 6, 74

		EMReadScreen stat_prog_cash_II_appl_date, 		8, 7, 33
		EMReadScreen stat_prog_cash_II_elig_begin_date, 8, 7, 44
		EMReadScreen stat_prog_cash_II_interview_date, 	8, 7, 55
		EMReadScreen stat_prog_cash_II_prog, 			2, 7, 67
		EMReadScreen stat_prog_cash_II_status, 			4, 7, 74
		EMReadScreen stat_prog_emer_appl_date, 			8, 8, 33
		' EMReadScreen stat_prog_emer_elig_begin_date, 	8, 8, 44
		EMReadScreen stat_prog_emer_interview_date, 	8, 8, 55
		EMReadScreen stat_prog_emer_prog, 				2, 8, 67
		EMReadScreen stat_prog_emer_status, 			4, 8, 74
		EMReadScreen stat_prog_grh_appl_date, 			8, 9, 33
		EMReadScreen stat_prog_grh_elig_begin_date, 	8, 9, 44
		EMReadScreen stat_prog_grh_interview_date, 		8, 9, 55
		EMReadScreen stat_prog_grh_status, 				4, 9, 74
		EMReadScreen stat_prog_snap_appl_date, 			8, 10, 33
		EMReadScreen stat_prog_snap_elig_begin_date, 	8, 10, 44
		EMReadScreen stat_prog_snap_interview_date, 	8, 10, 55
		EMReadScreen stat_prog_snap_status, 			4, 10, 74
		EMReadScreen stat_prog_ive_appl_date, 			8, 11, 33
		' EMReadScreen stat_prog_ive_elig_begin_date, 	8, 11, 44
		' EMReadScreen stat_prog_ive_interview_date, 		8, 11, 55
		EMReadScreen stat_prog_ive_status, 				4, 11, 74
		EMReadScreen stat_prog_hc_appl_date, 			8, 12, 33
		' EMReadScreen stat_prog_hc_elig_begin_date, 		8, 12, 44
		' EMReadScreen stat_prog_hc_interview_date, 		8, 12, 55
		EMReadScreen stat_prog_hc_status, 				4, 12, 74

		stat_prog_cash_I_appl_date = replace(stat_prog_cash_I_appl_date, " ", "/")
		If stat_prog_cash_I_appl_date = "__/__/__" Then stat_prog_cash_I_appl_date = ""

		stat_prog_cash_I_elig_begin_date = replace(stat_prog_cash_I_elig_begin_date, " ", "/")
		If stat_prog_cash_I_elig_begin_date = "__/__/__" Then stat_prog_cash_I_elig_begin_date = ""
		stat_prog_cash_I_interview_date = replace(stat_prog_cash_I_interview_date, " ", "/")
		If stat_prog_cash_I_interview_date = "__/__/__" Then stat_prog_cash_I_interview_date = ""
		stat_prog_cash_II_appl_date = replace(stat_prog_cash_II_appl_date, " ", "/")
		If stat_prog_cash_II_appl_date = "__/__/__" Then stat_prog_cash_II_appl_date = ""
		stat_prog_cash_II_elig_begin_date = replace(stat_prog_cash_II_elig_begin_date, " ", "/")
		If stat_prog_cash_II_elig_begin_date = "__/__/__" Then stat_prog_cash_II_elig_begin_date = ""
		stat_prog_cash_II_interview_date = replace(stat_prog_cash_II_interview_date, " ", "/")
		If stat_prog_cash_II_interview_date = "__/__/__" Then stat_prog_cash_II_interview_date = ""
		stat_prog_emer_appl_date = replace(stat_prog_emer_appl_date, " ", "/")
		If stat_prog_emer_appl_date = "__/__/__" Then stat_prog_emer_appl_date = ""
		stat_prog_emer_interview_date = replace(stat_prog_emer_interview_date, " ", "/")
		If stat_prog_emer_interview_date = "__/__/__" Then stat_prog_emer_interview_date = ""
		stat_prog_grh_appl_date = replace(stat_prog_grh_appl_date, " ", "/")
		If stat_prog_grh_appl_date = "__/__/__" Then stat_prog_grh_appl_date = ""
		stat_prog_grh_elig_begin_date = replace(stat_prog_grh_elig_begin_date, " ", "/")
		If stat_prog_grh_elig_begin_date = "__/__/__" Then stat_prog_grh_elig_begin_date = ""
		stat_prog_grh_interview_date = replace(stat_prog_grh_interview_date, " ", "/")
		If stat_prog_grh_interview_date = "__/__/__" Then stat_prog_grh_interview_date = ""
		stat_prog_snap_appl_date = replace(stat_prog_snap_appl_date, " ", "/")
		If stat_prog_snap_appl_date = "__/__/__" Then stat_prog_snap_appl_date = ""
		stat_prog_snap_elig_begin_date = replace(stat_prog_snap_elig_begin_date, " ", "/")
		If stat_prog_snap_elig_begin_date = "__/__/__" Then stat_prog_snap_elig_begin_date = ""
		stat_prog_snap_interview_date = replace(stat_prog_snap_interview_date, " ", "/")
		If stat_prog_snap_interview_date = "__/__/__" Then stat_prog_snap_interview_date = ""
		stat_prog_ive_appl_date = replace(stat_prog_ive_appl_date, " ", "/")
		If stat_prog_ive_appl_date = "__/__/__" Then stat_prog_ive_appl_date = ""
		stat_prog_hc_appl_date = replace(stat_prog_hc_appl_date, " ", "/")
		If stat_prog_hc_appl_date = "__/__/__" Then stat_prog_hc_appl_date = ""

		If stat_prog_cash_I_prog = "MF" Then stat_prog_cash_I_prog = "MFIP"
		If stat_prog_cash_I_prog = "RC" Then stat_prog_cash_I_prog = "RCA"
		If stat_prog_cash_I_prog = "MS" Then stat_prog_cash_I_prog = "MSA"
		If stat_prog_cash_I_prog = "GA" Then stat_prog_cash_I_prog = "GA"
		If stat_prog_cash_I_prog = "DW" Then stat_prog_cash_I_prog = "DWP"
		If stat_prog_cash_II_prog = "MF" Then stat_prog_cash_II_prog = "MFIP"
		If stat_prog_cash_II_prog = "RC" Then stat_prog_cash_II_prog = "RCA"
		If stat_prog_cash_II_prog = "MS" Then stat_prog_cash_II_prog = "MSA"
		If stat_prog_cash_II_prog = "GA" Then stat_prog_cash_II_prog = "GA"
		If stat_prog_cash_II_prog = "DW" Then stat_prog_cash_II_prog = "DWP"


		ReDim stat_memb_ref_numb(0)
		ReDim stat_memb_first_name(0)
		ReDim stat_memb_last_name(0)
		ReDim stat_memb_middle_initial(0)
		ReDim stat_memb_full_name(0)
		ReDim stat_memb_full_name_no_initial(0)
		ReDim stat_memb_full_name_last_name_first(0)
		ReDim stat_memb_full_name_last_name_first_no_mi(0)
		ReDim stat_memb_age(0)
		ReDim stat_memb_id_verif_code(0)
		ReDim stat_memb_id_verif_info(0)
		ReDim stat_memb_rel_to_applct_code(0)
		ReDim stat_memb_rel_to_applct_info(0)
		ReDim stat_memi_spouse_ref_numb(0)
		ReDim stat_memi_citizenship_yn(0)
		ReDim stat_memi_citizenship_verif_code(0)
		ReDim stat_memi_citizenship_verif_info(0)
		ReDim stat_jobs_one_exists(0)
		ReDim stat_jobs_one_job_ended(0)
		ReDim stat_jobs_one_job_counted(0)
		ReDim stat_jobs_one_inc_type(0)
		ReDim stat_jobs_one_sub_inc_type(0)
		ReDim stat_jobs_one_verif_code(0)
		ReDim stat_jobs_one_verif_info(0)
		ReDim stat_jobs_one_employer_name(0)
		ReDim stat_jobs_one_inc_start_date(0)
		ReDim stat_jobs_one_inc_end_date(0)
		ReDim stat_jobs_one_main_pay_freq(0)
		ReDim stat_jobs_one_snap_pic_pay_freq(0)
		ReDim stat_jobs_one_snap_pic_ave_hrs_per_pay(0)
		ReDim stat_jobs_one_snap_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_one_snap_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_one_grh_pic_pay_freq(0)
		ReDim stat_jobs_one_grh_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_one_grh_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_two_exists(0)
		ReDim stat_jobs_two_job_ended(0)
		ReDim stat_jobs_two_job_counted(0)
		ReDim stat_jobs_two_inc_type(0)
		ReDim stat_jobs_two_sub_inc_type(0)
		ReDim stat_jobs_two_verif_code(0)
		ReDim stat_jobs_two_verif_info(0)
		ReDim stat_jobs_two_employer_name(0)
		ReDim stat_jobs_two_inc_start_date(0)
		ReDim stat_jobs_two_inc_end_date(0)
		ReDim stat_jobs_two_main_pay_freq(0)
		ReDim stat_jobs_two_snap_pic_pay_freq(0)
		ReDim stat_jobs_two_snap_pic_ave_hrs_per_pay(0)
		ReDim stat_jobs_two_snap_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_two_snap_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_two_grh_pic_pay_freq(0)
		ReDim stat_jobs_two_grh_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_two_grh_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_three_exists(0)
		ReDim stat_jobs_three_job_ended(0)
		ReDim stat_jobs_three_job_counted(0)
		ReDim stat_jobs_three_inc_type(0)
		ReDim stat_jobs_three_sub_inc_type(0)
		ReDim stat_jobs_three_verif_code(0)
		ReDim stat_jobs_three_verif_info(0)
		ReDim stat_jobs_three_employer_name(0)
		ReDim stat_jobs_three_inc_start_date(0)
		ReDim stat_jobs_three_inc_end_date(0)
		ReDim stat_jobs_three_main_pay_freq(0)
		ReDim stat_jobs_three_snap_pic_pay_freq(0)
		ReDim stat_jobs_three_snap_pic_ave_hrs_per_pay(0)
		ReDim stat_jobs_three_snap_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_three_snap_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_three_grh_pic_pay_freq(0)
		ReDim stat_jobs_three_grh_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_three_grh_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_four_exists(0)
		ReDim stat_jobs_four_job_ended(0)
		ReDim stat_jobs_four_job_counted(0)
		ReDim stat_jobs_four_inc_type(0)
		ReDim stat_jobs_four_sub_inc_type(0)
		ReDim stat_jobs_four_verif_code(0)
		ReDim stat_jobs_four_verif_info(0)
		ReDim stat_jobs_four_employer_name(0)
		ReDim stat_jobs_four_inc_start_date(0)
		ReDim stat_jobs_four_inc_end_date(0)
		ReDim stat_jobs_four_main_pay_freq(0)
		ReDim stat_jobs_four_snap_pic_pay_freq(0)
		ReDim stat_jobs_four_snap_pic_ave_hrs_per_pay(0)
		ReDim stat_jobs_four_snap_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_four_snap_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_four_grh_pic_pay_freq(0)
		ReDim stat_jobs_four_grh_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_four_grh_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_five_exists(0)
		ReDim stat_jobs_five_job_ended(0)
		ReDim stat_jobs_five_job_counted(0)
		ReDim stat_jobs_five_inc_type(0)
		ReDim stat_jobs_five_sub_inc_type(0)
		ReDim stat_jobs_five_verif_code(0)
		ReDim stat_jobs_five_verif_info(0)
		ReDim stat_jobs_five_employer_name(0)
		ReDim stat_jobs_five_inc_start_date(0)
		ReDim stat_jobs_five_inc_end_date(0)
		ReDim stat_jobs_five_main_pay_freq(0)
		ReDim stat_jobs_five_snap_pic_pay_freq(0)
		ReDim stat_jobs_five_snap_pic_ave_hrs_per_pay(0)
		ReDim stat_jobs_five_snap_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_five_snap_pic_prosp_monthly_inc(0)
		ReDim stat_jobs_five_grh_pic_pay_freq(0)
		ReDim stat_jobs_five_grh_pic_ave_inc_per_pay(0)
		ReDim stat_jobs_five_grh_pic_prosp_monthly_inc(0)
		ReDim stat_busi_one_exists(0)
		ReDim stat_busi_one_type(0)
		ReDim stat_busi_one_counted(0)
		ReDim stat_busi_one_type_info(0)
		ReDim stat_busi_one_inc_start_date(0)
		ReDim stat_busi_one_inc_end_date(0)
		ReDim stat_busi_one_method(0)
		ReDim stat_busi_one_method_date(0)
		ReDim stat_busi_one_snap_retro_net_inc(0)
		ReDim stat_busi_one_snap_prosp_net_inc(0)
		ReDim stat_busi_one_snap_retro_gross_inc(0)
		ReDim stat_busi_one_snap_retro_expenses(0)
		ReDim stat_busi_one_snap_income_verif_code(0)
		ReDim stat_busi_one_snap_income_verif_info(0)
		ReDim stat_busi_one_snap_prosp_gross_inc(0)
		ReDim stat_busi_one_snap_prosp_expenses(0)
		ReDim stat_busi_one_snap_expense_verif_code(0)
		ReDim stat_busi_one_snap_expense_verif_info(0)
		ReDim stat_busi_two_exists(0)
		ReDim stat_busi_two_type(0)
		ReDim stat_busi_two_counted(0)
		ReDim stat_busi_two_type_info(0)
		ReDim stat_busi_two_inc_start_date(0)
		ReDim stat_busi_two_inc_end_date(0)
		ReDim stat_busi_two_method(0)
		ReDim stat_busi_two_method_date(0)
		ReDim stat_busi_two_snap_retro_net_inc(0)
		ReDim stat_busi_two_snap_prosp_net_inc(0)
		ReDim stat_busi_two_snap_retro_gross_inc(0)
		ReDim stat_busi_two_snap_retro_expenses(0)
		ReDim stat_busi_two_snap_income_verif_code(0)
		ReDim stat_busi_two_snap_income_verif_info(0)
		ReDim stat_busi_two_snap_prosp_gross_inc(0)
		ReDim stat_busi_two_snap_prosp_expenses(0)
		ReDim stat_busi_two_snap_expense_verif_code(0)
		ReDim stat_busi_two_snap_expense_verif_info(0)
		ReDim stat_busi_three_exists(0)
		ReDim stat_busi_three_type(0)
		ReDim stat_busi_three_counted(0)
		ReDim stat_busi_three_type_info(0)
		ReDim stat_busi_three_inc_start_date(0)
		ReDim stat_busi_three_inc_end_date(0)
		ReDim stat_busi_three_method(0)
		ReDim stat_busi_three_method_date(0)
		ReDim stat_busi_three_snap_retro_net_inc(0)
		ReDim stat_busi_three_snap_prosp_net_inc(0)
		ReDim stat_busi_three_snap_retro_gross_inc(0)
		ReDim stat_busi_three_snap_retro_expenses(0)
		ReDim stat_busi_three_snap_income_verif_code(0)
		ReDim stat_busi_three_snap_income_verif_info(0)
		ReDim stat_busi_three_snap_prosp_gross_inc(0)
		ReDim stat_busi_three_snap_prosp_expenses(0)
		ReDim stat_busi_three_snap_expense_verif_code(0)
		ReDim stat_busi_three_snap_expense_verif_info(0)
		ReDim stat_unea_one_exists(0)
		ReDim stat_unea_one_counted(0)
		ReDim stat_unea_one_type_code(0)
		ReDim stat_unea_one_type_info(0)
		ReDim stat_unea_one_verif_code(0)
		ReDim stat_unea_one_verif_info(0)
		ReDim stat_unea_one_inc_start_date(0)
		ReDim stat_unea_one_inc_end_date(0)
		ReDim stat_unea_one_snap_pic_pay_freq(0)
		ReDim stat_unea_one_snap_pic_ave_inc_per_pay(0)
		ReDim stat_unea_one_snap_pic_prosp_monthly_inc(0)
		ReDim stat_unea_two_exists(0)
		ReDim stat_unea_two_counted(0)
		ReDim stat_unea_two_type_code(0)
		ReDim stat_unea_two_type_info(0)
		ReDim stat_unea_two_verif_code(0)
		ReDim stat_unea_two_verif_info(0)
		ReDim stat_unea_two_inc_start_date(0)
		ReDim stat_unea_two_inc_end_date(0)
		ReDim stat_unea_two_snap_pic_pay_freq(0)
		ReDim stat_unea_two_snap_pic_ave_inc_per_pay(0)
		ReDim stat_unea_two_snap_pic_prosp_monthly_inc(0)
		ReDim stat_unea_three_exists(0)
		ReDim stat_unea_three_counted(0)
		ReDim stat_unea_three_type_code(0)
		ReDim stat_unea_three_type_info(0)
		ReDim stat_unea_three_verif_code(0)
		ReDim stat_unea_three_verif_info(0)
		ReDim stat_unea_three_inc_start_date(0)
		ReDim stat_unea_three_inc_end_date(0)
		ReDim stat_unea_three_snap_pic_pay_freq(0)
		ReDim stat_unea_three_snap_pic_ave_inc_per_pay(0)
		ReDim stat_unea_three_snap_pic_prosp_monthly_inc(0)
		ReDim stat_unea_four_exists(0)
		ReDim stat_unea_four_counted(0)
		ReDim stat_unea_four_type_code(0)
		ReDim stat_unea_four_type_info(0)
		ReDim stat_unea_four_verif_code(0)
		ReDim stat_unea_four_verif_info(0)
		ReDim stat_unea_four_inc_start_date(0)
		ReDim stat_unea_four_inc_end_date(0)
		ReDim stat_unea_four_snap_pic_pay_freq(0)
		ReDim stat_unea_four_snap_pic_ave_inc_per_pay(0)
		ReDim stat_unea_four_snap_pic_prosp_monthly_inc(0)
		ReDim stat_unea_five_exists(0)
		ReDim stat_unea_five_counted(0)
		ReDim stat_unea_five_type_code(0)
		ReDim stat_unea_five_type_info(0)
		ReDim stat_unea_five_verif_code(0)
		ReDim stat_unea_five_verif_info(0)
		ReDim stat_unea_five_inc_start_date(0)
		ReDim stat_unea_five_inc_end_date(0)
		ReDim stat_unea_five_snap_pic_pay_freq(0)
		ReDim stat_unea_five_snap_pic_ave_inc_per_pay(0)
		ReDim stat_unea_five_snap_pic_prosp_monthly_inc(0)
		ReDim stat_acct_one_exists(0)
		ReDim stat_acct_one_type(0)
		ReDim stat_acct_one_balence(0)
		ReDim stat_acct_one_count_snap_yn(0)
		ReDim stat_acct_two_exists(0)
		ReDim stat_acct_two_type(0)
		ReDim stat_acct_two_balence(0)
		ReDim stat_acct_two_count_snap_yn(0)
		ReDim stat_acct_three_exists(0)
		ReDim stat_acct_three_type(0)
		ReDim stat_acct_three_balence(0)
		ReDim stat_acct_three_count_snap_yn(0)
		ReDim stat_acct_four_exists(0)
		ReDim stat_acct_four_type(0)
		ReDim stat_acct_four_balence(0)
		ReDim stat_acct_four_count_snap_yn(0)
		ReDim stat_acct_five_exists(0)
		ReDim stat_acct_five_type(0)
		ReDim stat_acct_five_balence(0)
		ReDim stat_acct_five_count_snap_yn(0)
		ReDim stat_shel_exists(0)
		ReDim stat_shel_subsidized_yn(0)
		ReDim stat_shel_shared_yn(0)
		ReDim stat_shel_paid_to(0)
		ReDim stat_shel_retro_rent_amount(0)
		ReDim stat_shel_retro_rent_verif_code(0)
		ReDim stat_shel_retro_rent_verif_info(0)
		ReDim stat_shel_prosp_rent_amount(0)
		ReDim stat_shel_prosp_rent_verif_code(0)
		ReDim stat_shel_prosp_rent_verif_info(0)
		ReDim stat_shel_retro_lot_rent_amount(0)
		ReDim stat_shel_retro_lot_rent_verif_code(0)
		ReDim stat_shel_retro_lot_rent_verif_info(0)
		ReDim stat_shel_prosp_lot_rent_amount(0)
		ReDim stat_shel_prosp_lot_rent_verif_code(0)
		ReDim stat_shel_prosp_lot_rent_verif_info(0)
		ReDim stat_shel_retro_mortgage_amount(0)
		ReDim stat_shel_retro_mortgage_verif_code(0)
		ReDim stat_shel_retro_mortgage_verif_info(0)
		ReDim stat_shel_prosp_mortgage_amount(0)
		ReDim stat_shel_prosp_mortgage_verif_code(0)
		ReDim stat_shel_prosp_mortgage_verif_info(0)
		ReDim stat_shel_retro_insurance_amount(0)
		ReDim stat_shel_retro_insurance_verif_code(0)
		ReDim stat_shel_retro_insurance_verif_info(0)
		ReDim stat_shel_prosp_insurance_amount(0)
		ReDim stat_shel_prosp_insurance_verif_code(0)
		ReDim stat_shel_prosp_insurance_verif_info(0)
		ReDim stat_shel_retro_taxes_amount(0)
		ReDim stat_shel_retro_taxes_verif_code(0)
		ReDim stat_shel_retro_taxes_verif_info(0)
		ReDim stat_shel_prosp_taxes_amount(0)
		ReDim stat_shel_prosp_taxes_verif_code(0)
		ReDim stat_shel_prosp_taxes_verif_info(0)
		ReDim stat_shel_retro_room_amount(0)
		ReDim stat_shel_retro_room_verif_code(0)
		ReDim stat_shel_retro_room_verif_info(0)
		ReDim stat_shel_prosp_room_amount(0)
		ReDim stat_shel_prosp_room_verif_code(0)
		ReDim stat_shel_prosp_room_verif_info(0)
		ReDim stat_shel_retro_garage_amount(0)
		ReDim stat_shel_retro_garage_verif_code(0)
		ReDim stat_shel_retro_garage_verif_info(0)
		ReDim stat_shel_prosp_garage_amount(0)
		ReDim stat_shel_prosp_garage_verif_code(0)
		ReDim stat_shel_prosp_garage_verif_info(0)
		ReDim stat_shel_retro_subsidy_amount(0)
		ReDim stat_shel_retro_subsidy_verif_code(0)
		ReDim stat_shel_retro_subsidy_verif_info(0)
		ReDim stat_shel_prosp_subsidy_amount(0)
		ReDim stat_shel_prosp_subsidy_verif_code(0)
		ReDim stat_shel_prosp_subsidy_verif_info(0)

		ReDim stat_disq_one_exists(0)
		ReDim stat_disq_one_program(0)
		ReDim stat_disq_one_type_code(0)
		ReDim stat_disq_one_type_info(0)
		ReDim stat_disq_one_begin_date(0)
		ReDim stat_disq_one_end_date(0)
		ReDim stat_disq_one_cure_reason_code(0)
		ReDim stat_disq_one_cure_reason_info(0)
		ReDim stat_disq_one_fraud_determination_date(0)
		ReDim stat_disq_one_county_of_fraud(0)
		ReDim stat_disq_one_state_of_fraud(0)
		ReDim stat_disq_one_SNAP_trafficking_yn(0)
		ReDim stat_disq_one_SNAP_offense_code(0)
		ReDim stat_disq_one_SNAP_offense_info(0)
		ReDim stat_disq_one_source(0)
		ReDim stat_disq_one_active(0)

		ReDim stat_disq_two_exists(0)
		ReDim stat_disq_two_program(0)
		ReDim stat_disq_two_type_code(0)
		ReDim stat_disq_two_type_info(0)
		ReDim stat_disq_two_begin_date(0)
		ReDim stat_disq_two_end_date(0)
		ReDim stat_disq_two_cure_reason_code(0)
		ReDim stat_disq_two_cure_reason_info(0)
		ReDim stat_disq_two_fraud_determination_date(0)
		ReDim stat_disq_two_county_of_fraud(0)
		ReDim stat_disq_two_state_of_fraud(0)
		ReDim stat_disq_two_SNAP_trafficking_yn(0)
		ReDim stat_disq_two_SNAP_offense_code(0)
		ReDim stat_disq_two_SNAP_offense_info(0)
		ReDim stat_disq_two_source(0)
		ReDim stat_disq_two_active(0)

		ReDim stat_disq_three_exists(0)
		ReDim stat_disq_three_program(0)
		ReDim stat_disq_three_type_code(0)
		ReDim stat_disq_three_type_info(0)
		ReDim stat_disq_three_begin_date(0)
		ReDim stat_disq_three_end_date(0)
		ReDim stat_disq_three_cure_reason_code(0)
		ReDim stat_disq_three_cure_reason_info(0)
		ReDim stat_disq_three_fraud_determination_date(0)
		ReDim stat_disq_three_county_of_fraud(0)
		ReDim stat_disq_three_state_of_fraud(0)
		ReDim stat_disq_three_SNAP_trafficking_yn(0)
		ReDim stat_disq_three_SNAP_offense_code(0)
		ReDim stat_disq_three_SNAP_offense_info(0)
		ReDim stat_disq_three_source(0)
		ReDim stat_disq_three_active(0)

		ReDim stat_disq_four_exists(0)
		ReDim stat_disq_four_program(0)
		ReDim stat_disq_four_type_code(0)
		ReDim stat_disq_four_type_info(0)
		ReDim stat_disq_four_begin_date(0)
		ReDim stat_disq_four_end_date(0)
		ReDim stat_disq_four_cure_reason_code(0)
		ReDim stat_disq_four_cure_reason_info(0)
		ReDim stat_disq_four_fraud_determination_date(0)
		ReDim stat_disq_four_county_of_fraud(0)
		ReDim stat_disq_four_state_of_fraud(0)
		ReDim stat_disq_four_SNAP_trafficking_yn(0)
		ReDim stat_disq_four_SNAP_offense_code(0)
		ReDim stat_disq_four_SNAP_offense_info(0)
		ReDim stat_disq_four_source(0)
		ReDim stat_disq_four_active(0)

		ReDim stat_disq_five_exists(0)
		ReDim stat_disq_five_program(0)
		ReDim stat_disq_five_type_code(0)
		ReDim stat_disq_five_type_info(0)
		ReDim stat_disq_five_begin_date(0)
		ReDim stat_disq_five_end_date(0)
		ReDim stat_disq_five_cure_reason_code(0)
		ReDim stat_disq_five_cure_reason_info(0)
		ReDim stat_disq_five_fraud_determination_date(0)
		ReDim stat_disq_five_county_of_fraud(0)
		ReDim stat_disq_five_state_of_fraud(0)
		ReDim stat_disq_five_SNAP_trafficking_yn(0)
		ReDim stat_disq_five_SNAP_offense_code(0)
		ReDim stat_disq_five_SNAP_offense_info(0)
		ReDim stat_disq_five_source(0)
		ReDim stat_disq_five_active(0)

		stat_shel_prosp_all_total = 0

		Call navigate_to_MAXIS_screen("STAT", "MEMB")
		memb_count = -1
		Do
			memb_count = memb_count + 1

			ReDim preserve stat_memb_ref_numb(memb_count)
			ReDim preserve stat_memb_first_name(memb_count)
			ReDim preserve stat_memb_last_name(memb_count)
			ReDim preserve stat_memb_middle_initial(memb_count)
			ReDim preserve stat_memb_full_name(memb_count)
			ReDim preserve stat_memb_full_name_no_initial(memb_count)
			ReDim preserve stat_memb_full_name_last_name_first(memb_count)
			ReDim preserve stat_memb_full_name_last_name_first_no_mi(memb_count)
			ReDim preserve stat_memb_age(memb_count)
			ReDim preserve stat_memb_id_verif_code(memb_count)
			ReDim preserve stat_memb_id_verif_info(memb_count)
			ReDim preserve stat_memb_rel_to_applct_code(memb_count)
			ReDim preserve stat_memb_rel_to_applct_info(memb_count)
			ReDim preserve stat_memi_spouse_ref_numb(memb_count)
			ReDim preserve stat_memi_citizenship_yn(memb_count)
			ReDim preserve stat_memi_citizenship_verif_code(memb_count)
			ReDim preserve stat_memi_citizenship_verif_info(memb_count)
			ReDim preserve stat_jobs_one_exists(memb_count)
			ReDim preserve stat_jobs_one_job_ended(memb_count)
			ReDim preserve stat_jobs_one_job_counted(memb_count)
			ReDim preserve stat_jobs_one_inc_type(memb_count)
			ReDim preserve stat_jobs_one_sub_inc_type(memb_count)
			ReDim preserve stat_jobs_one_verif_code(memb_count)
			ReDim preserve stat_jobs_one_verif_info(memb_count)
			ReDim preserve stat_jobs_one_employer_name(memb_count)
			ReDim preserve stat_jobs_one_inc_start_date(memb_count)
			ReDim preserve stat_jobs_one_inc_end_date(memb_count)
			ReDim preserve stat_jobs_one_main_pay_freq(memb_count)
			ReDim preserve stat_jobs_one_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_one_snap_pic_ave_hrs_per_pay(memb_count)
			ReDim preserve stat_jobs_one_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_one_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_one_grh_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_one_grh_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_one_grh_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_two_exists(memb_count)
			ReDim preserve stat_jobs_two_job_ended(memb_count)
			ReDim preserve stat_jobs_two_job_counted(memb_count)
			ReDim preserve stat_jobs_two_inc_type(memb_count)
			ReDim preserve stat_jobs_two_sub_inc_type(memb_count)
			ReDim preserve stat_jobs_two_verif_code(memb_count)
			ReDim preserve stat_jobs_two_verif_info(memb_count)
			ReDim preserve stat_jobs_two_employer_name(memb_count)
			ReDim preserve stat_jobs_two_inc_start_date(memb_count)
			ReDim preserve stat_jobs_two_inc_end_date(memb_count)
			ReDim preserve stat_jobs_two_main_pay_freq(memb_count)
			ReDim preserve stat_jobs_two_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_two_snap_pic_ave_hrs_per_pay(memb_count)
			ReDim preserve stat_jobs_two_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_two_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_two_grh_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_two_grh_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_two_grh_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_three_exists(memb_count)
			ReDim preserve stat_jobs_three_job_ended(memb_count)
			ReDim preserve stat_jobs_three_job_counted(memb_count)
			ReDim preserve stat_jobs_three_inc_type(memb_count)
			ReDim preserve stat_jobs_three_sub_inc_type(memb_count)
			ReDim preserve stat_jobs_three_verif_code(memb_count)
			ReDim preserve stat_jobs_three_verif_info(memb_count)
			ReDim preserve stat_jobs_three_employer_name(memb_count)
			ReDim preserve stat_jobs_three_inc_start_date(memb_count)
			ReDim preserve stat_jobs_three_inc_end_date(memb_count)
			ReDim preserve stat_jobs_three_main_pay_freq(memb_count)
			ReDim preserve stat_jobs_three_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_three_snap_pic_ave_hrs_per_pay(memb_count)
			ReDim preserve stat_jobs_three_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_three_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_three_grh_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_three_grh_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_three_grh_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_four_exists(memb_count)
			ReDim preserve stat_jobs_four_job_ended(memb_count)
			ReDim preserve stat_jobs_four_job_counted(memb_count)
			ReDim preserve stat_jobs_four_inc_type(memb_count)
			ReDim preserve stat_jobs_four_sub_inc_type(memb_count)
			ReDim preserve stat_jobs_four_verif_code(memb_count)
			ReDim preserve stat_jobs_four_verif_info(memb_count)
			ReDim preserve stat_jobs_four_employer_name(memb_count)
			ReDim preserve stat_jobs_four_inc_start_date(memb_count)
			ReDim preserve stat_jobs_four_inc_end_date(memb_count)
			ReDim preserve stat_jobs_four_main_pay_freq(memb_count)
			ReDim preserve stat_jobs_four_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_four_snap_pic_ave_hrs_per_pay(memb_count)
			ReDim preserve stat_jobs_four_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_four_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_four_grh_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_four_grh_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_four_grh_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_five_exists(memb_count)
			ReDim preserve stat_jobs_five_job_ended(memb_count)
			ReDim preserve stat_jobs_five_job_counted(memb_count)
			ReDim preserve stat_jobs_five_inc_type(memb_count)
			ReDim preserve stat_jobs_five_sub_inc_type(memb_count)
			ReDim preserve stat_jobs_five_verif_code(memb_count)
			ReDim preserve stat_jobs_five_verif_info(memb_count)
			ReDim preserve stat_jobs_five_employer_name(memb_count)
			ReDim preserve stat_jobs_five_inc_start_date(memb_count)
			ReDim preserve stat_jobs_five_inc_end_date(memb_count)
			ReDim preserve stat_jobs_five_main_pay_freq(memb_count)
			ReDim preserve stat_jobs_five_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_five_snap_pic_ave_hrs_per_pay(memb_count)
			ReDim preserve stat_jobs_five_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_five_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_jobs_five_grh_pic_pay_freq(memb_count)
			ReDim preserve stat_jobs_five_grh_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_jobs_five_grh_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_busi_one_exists(memb_count)
			ReDim preserve stat_busi_one_type(memb_count)
			ReDim preserve stat_busi_one_counted(memb_count)
			ReDim preserve stat_busi_one_type_info(memb_count)
			ReDim preserve stat_busi_one_inc_start_date(memb_count)
			ReDim preserve stat_busi_one_inc_end_date(memb_count)
			ReDim preserve stat_busi_one_method(memb_count)
			ReDim preserve stat_busi_one_method_date(memb_count)
			ReDim preserve stat_busi_one_snap_retro_net_inc(memb_count)
			ReDim preserve stat_busi_one_snap_prosp_net_inc(memb_count)
			ReDim preserve stat_busi_one_snap_retro_gross_inc(memb_count)
			ReDim preserve stat_busi_one_snap_retro_expenses(memb_count)
			ReDim preserve stat_busi_one_snap_income_verif_code(memb_count)
			ReDim preserve stat_busi_one_snap_income_verif_info(memb_count)
			ReDim preserve stat_busi_one_snap_prosp_gross_inc(memb_count)
			ReDim preserve stat_busi_one_snap_prosp_expenses(memb_count)
			ReDim preserve stat_busi_one_snap_expense_verif_code(memb_count)
			ReDim preserve stat_busi_one_snap_expense_verif_info(memb_count)
			ReDim preserve stat_busi_two_exists(memb_count)
			ReDim preserve stat_busi_two_type(memb_count)
			ReDim preserve stat_busi_two_counted(memb_count)
			ReDim preserve stat_busi_two_type_info(memb_count)
			ReDim preserve stat_busi_two_inc_start_date(memb_count)
			ReDim preserve stat_busi_two_inc_end_date(memb_count)
			ReDim preserve stat_busi_two_method(memb_count)
			ReDim preserve stat_busi_two_method_date(memb_count)
			ReDim preserve stat_busi_two_snap_retro_net_inc(memb_count)
			ReDim preserve stat_busi_two_snap_prosp_net_inc(memb_count)
			ReDim preserve stat_busi_two_snap_retro_gross_inc(memb_count)
			ReDim preserve stat_busi_two_snap_retro_expenses(memb_count)
			ReDim preserve stat_busi_two_snap_income_verif_code(memb_count)
			ReDim preserve stat_busi_two_snap_income_verif_info(memb_count)
			ReDim preserve stat_busi_two_snap_prosp_gross_inc(memb_count)
			ReDim preserve stat_busi_two_snap_prosp_expenses(memb_count)
			ReDim preserve stat_busi_two_snap_expense_verif_code(memb_count)
			ReDim preserve stat_busi_two_snap_expense_verif_info(memb_count)
			ReDim preserve stat_busi_three_exists(memb_count)
			ReDim preserve stat_busi_three_type(memb_count)
			ReDim preserve stat_busi_three_counted(memb_count)
			ReDim preserve stat_busi_three_type_info(memb_count)
			ReDim preserve stat_busi_three_inc_start_date(memb_count)
			ReDim preserve stat_busi_three_inc_end_date(memb_count)
			ReDim preserve stat_busi_three_method(memb_count)
			ReDim preserve stat_busi_three_method_date(memb_count)
			ReDim preserve stat_busi_three_snap_retro_net_inc(memb_count)
			ReDim preserve stat_busi_three_snap_prosp_net_inc(memb_count)
			ReDim preserve stat_busi_three_snap_retro_gross_inc(memb_count)
			ReDim preserve stat_busi_three_snap_retro_expenses(memb_count)
			ReDim preserve stat_busi_three_snap_income_verif_code(memb_count)
			ReDim preserve stat_busi_three_snap_income_verif_info(memb_count)
			ReDim preserve stat_busi_three_snap_prosp_gross_inc(memb_count)
			ReDim preserve stat_busi_three_snap_prosp_expenses(memb_count)
			ReDim preserve stat_busi_three_snap_expense_verif_code(memb_count)
			ReDim preserve stat_busi_three_snap_expense_verif_info(memb_count)
			ReDim preserve stat_unea_one_exists(memb_count)
			ReDim preserve stat_unea_one_counted(memb_count)
			ReDim preserve stat_unea_one_type_code(memb_count)
			ReDim preserve stat_unea_one_type_info(memb_count)
			ReDim preserve stat_unea_one_verif_code(memb_count)
			ReDim preserve stat_unea_one_verif_info(memb_count)
			ReDim preserve stat_unea_one_inc_start_date(memb_count)
			ReDim preserve stat_unea_one_inc_end_date(memb_count)
			ReDim preserve stat_unea_one_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_unea_one_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_unea_one_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_unea_two_exists(memb_count)
			ReDim preserve stat_unea_two_counted(memb_count)
			ReDim preserve stat_unea_two_type_code(memb_count)
			ReDim preserve stat_unea_two_type_info(memb_count)
			ReDim preserve stat_unea_two_verif_code(memb_count)
			ReDim preserve stat_unea_two_verif_info(memb_count)
			ReDim preserve stat_unea_two_inc_start_date(memb_count)
			ReDim preserve stat_unea_two_inc_end_date(memb_count)
			ReDim preserve stat_unea_two_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_unea_two_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_unea_two_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_unea_three_exists(memb_count)
			ReDim preserve stat_unea_three_counted(memb_count)
			ReDim preserve stat_unea_three_type_code(memb_count)
			ReDim preserve stat_unea_three_type_info(memb_count)
			ReDim preserve stat_unea_three_verif_code(memb_count)
			ReDim preserve stat_unea_three_verif_info(memb_count)
			ReDim preserve stat_unea_three_inc_start_date(memb_count)
			ReDim preserve stat_unea_three_inc_end_date(memb_count)
			ReDim preserve stat_unea_three_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_unea_three_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_unea_three_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_unea_four_exists(memb_count)
			ReDim preserve stat_unea_four_counted(memb_count)
			ReDim preserve stat_unea_four_type_code(memb_count)
			ReDim preserve stat_unea_four_type_info(memb_count)
			ReDim preserve stat_unea_four_verif_code(memb_count)
			ReDim preserve stat_unea_four_verif_info(memb_count)
			ReDim preserve stat_unea_four_inc_start_date(memb_count)
			ReDim preserve stat_unea_four_inc_end_date(memb_count)
			ReDim preserve stat_unea_four_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_unea_four_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_unea_four_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_unea_five_exists(memb_count)
			ReDim preserve stat_unea_five_counted(memb_count)
			ReDim preserve stat_unea_five_type_code(memb_count)
			ReDim preserve stat_unea_five_type_info(memb_count)
			ReDim preserve stat_unea_five_verif_code(memb_count)
			ReDim preserve stat_unea_five_verif_info(memb_count)
			ReDim preserve stat_unea_five_inc_start_date(memb_count)
			ReDim preserve stat_unea_five_inc_end_date(memb_count)
			ReDim preserve stat_unea_five_snap_pic_pay_freq(memb_count)
			ReDim preserve stat_unea_five_snap_pic_ave_inc_per_pay(memb_count)
			ReDim preserve stat_unea_five_snap_pic_prosp_monthly_inc(memb_count)
			ReDim preserve stat_acct_one_exists(memb_count)
			ReDim preserve stat_acct_one_type(memb_count)
			ReDim preserve stat_acct_one_balence(memb_count)
			ReDim preserve stat_acct_one_count_snap_yn(memb_count)
			ReDim preserve stat_acct_two_exists(memb_count)
			ReDim preserve stat_acct_two_type(memb_count)
			ReDim preserve stat_acct_two_balence(memb_count)
			ReDim preserve stat_acct_two_count_snap_yn(memb_count)
			ReDim preserve stat_acct_three_exists(memb_count)
			ReDim preserve stat_acct_three_type(memb_count)
			ReDim preserve stat_acct_three_balence(memb_count)
			ReDim preserve stat_acct_three_count_snap_yn(memb_count)
			ReDim preserve stat_acct_four_exists(memb_count)
			ReDim preserve stat_acct_four_type(memb_count)
			ReDim preserve stat_acct_four_balence(memb_count)
			ReDim preserve stat_acct_four_count_snap_yn(memb_count)
			ReDim preserve stat_acct_five_exists(memb_count)
			ReDim preserve stat_acct_five_type(memb_count)
			ReDim preserve stat_acct_five_balence(memb_count)
			ReDim preserve stat_acct_five_count_snap_yn(memb_count)
			ReDim preserve stat_shel_exists(memb_count)
			ReDim preserve stat_shel_subsidized_yn(memb_count)
			ReDim preserve stat_shel_shared_yn(memb_count)
			ReDim preserve stat_shel_paid_to(memb_count)
			ReDim preserve stat_shel_retro_rent_amount(memb_count)
			ReDim preserve stat_shel_retro_rent_verif_code(memb_count)
			ReDim preserve stat_shel_retro_rent_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_rent_amount(memb_count)
			ReDim preserve stat_shel_prosp_rent_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_rent_verif_info(memb_count)
			ReDim preserve stat_shel_retro_lot_rent_amount(memb_count)
			ReDim preserve stat_shel_retro_lot_rent_verif_code(memb_count)
			ReDim preserve stat_shel_retro_lot_rent_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_lot_rent_amount(memb_count)
			ReDim preserve stat_shel_prosp_lot_rent_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_lot_rent_verif_info(memb_count)
			ReDim preserve stat_shel_retro_mortgage_amount(memb_count)
			ReDim preserve stat_shel_retro_mortgage_verif_code(memb_count)
			ReDim preserve stat_shel_retro_mortgage_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_mortgage_amount(memb_count)
			ReDim preserve stat_shel_prosp_mortgage_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_mortgage_verif_info(memb_count)
			ReDim preserve stat_shel_retro_insurance_amount(memb_count)
			ReDim preserve stat_shel_retro_insurance_verif_code(memb_count)
			ReDim preserve stat_shel_retro_insurance_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_insurance_amount(memb_count)
			ReDim preserve stat_shel_prosp_insurance_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_insurance_verif_info(memb_count)
			ReDim preserve stat_shel_retro_taxes_amount(memb_count)
			ReDim preserve stat_shel_retro_taxes_verif_code(memb_count)
			ReDim preserve stat_shel_retro_taxes_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_taxes_amount(memb_count)
			ReDim preserve stat_shel_prosp_taxes_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_taxes_verif_info(memb_count)
			ReDim preserve stat_shel_retro_room_amount(memb_count)
			ReDim preserve stat_shel_retro_room_verif_code(memb_count)
			ReDim preserve stat_shel_retro_room_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_room_amount(memb_count)
			ReDim preserve stat_shel_prosp_room_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_room_verif_info(memb_count)
			ReDim preserve stat_shel_retro_garage_amount(memb_count)
			ReDim preserve stat_shel_retro_garage_verif_code(memb_count)
			ReDim preserve stat_shel_retro_garage_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_garage_amount(memb_count)
			ReDim preserve stat_shel_prosp_garage_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_garage_verif_info(memb_count)
			ReDim preserve stat_shel_retro_subsidy_amount(memb_count)
			ReDim preserve stat_shel_retro_subsidy_verif_code(memb_count)
			ReDim preserve stat_shel_retro_subsidy_verif_info(memb_count)
			ReDim preserve stat_shel_prosp_subsidy_amount(memb_count)
			ReDim preserve stat_shel_prosp_subsidy_verif_code(memb_count)
			ReDim preserve stat_shel_prosp_subsidy_verif_info(memb_count)

			ReDim preserve stat_disq_one_exists(memb_count)
			ReDim preserve stat_disq_one_program(memb_count)
			ReDim preserve stat_disq_one_type_code(memb_count)
			ReDim preserve stat_disq_one_type_info(memb_count)
			ReDim preserve stat_disq_one_begin_date(memb_count)
			ReDim preserve stat_disq_one_end_date(memb_count)
			ReDim preserve stat_disq_one_cure_reason_code(memb_count)
			ReDim preserve stat_disq_one_cure_reason_info(memb_count)
			ReDim preserve stat_disq_one_fraud_determination_date(memb_count)
			ReDim preserve stat_disq_one_county_of_fraud(memb_count)
			ReDim preserve stat_disq_one_state_of_fraud(memb_count)
			ReDim preserve stat_disq_one_SNAP_trafficking_yn(memb_count)
			ReDim preserve stat_disq_one_SNAP_offense_code(memb_count)
			ReDim preserve stat_disq_one_SNAP_offense_info(memb_count)
			ReDim preserve stat_disq_one_source(memb_count)
			ReDim preserve stat_disq_one_active(memb_count)


			ReDim preserve stat_disq_two_exists(memb_count)
			ReDim preserve stat_disq_two_program(memb_count)
			ReDim preserve stat_disq_two_type_code(memb_count)
			ReDim preserve stat_disq_two_type_info(memb_count)
			ReDim preserve stat_disq_two_begin_date(memb_count)
			ReDim preserve stat_disq_two_end_date(memb_count)
			ReDim preserve stat_disq_two_cure_reason_code(memb_count)
			ReDim preserve stat_disq_two_cure_reason_info(memb_count)
			ReDim preserve stat_disq_two_fraud_determination_date(memb_count)
			ReDim preserve stat_disq_two_county_of_fraud(memb_count)
			ReDim preserve stat_disq_two_state_of_fraud(memb_count)
			ReDim preserve stat_disq_two_SNAP_trafficking_yn(memb_count)
			ReDim preserve stat_disq_two_SNAP_offense_code(memb_count)
			ReDim preserve stat_disq_two_SNAP_offense_info(memb_count)
			ReDim preserve stat_disq_two_source(memb_count)
			ReDim preserve stat_disq_two_active(memb_count)

			ReDim preserve stat_disq_three_exists(memb_count)
			ReDim preserve stat_disq_three_program(memb_count)
			ReDim preserve stat_disq_three_type_code(memb_count)
			ReDim preserve stat_disq_three_type_info(memb_count)
			ReDim preserve stat_disq_three_begin_date(memb_count)
			ReDim preserve stat_disq_three_end_date(memb_count)
			ReDim preserve stat_disq_three_cure_reason_code(memb_count)
			ReDim preserve stat_disq_three_cure_reason_info(memb_count)
			ReDim preserve stat_disq_three_fraud_determination_date(memb_count)
			ReDim preserve stat_disq_three_county_of_fraud(memb_count)
			ReDim preserve stat_disq_three_state_of_fraud(memb_count)
			ReDim preserve stat_disq_three_SNAP_trafficking_yn(memb_count)
			ReDim preserve stat_disq_three_SNAP_offense_code(memb_count)
			ReDim preserve stat_disq_three_SNAP_offense_info(memb_count)
			ReDim preserve stat_disq_three_source(memb_count)
			ReDim preserve stat_disq_three_active(memb_count)

			ReDim preserve stat_disq_four_exists(memb_count)
			ReDim preserve stat_disq_four_program(memb_count)
			ReDim preserve stat_disq_four_type_code(memb_count)
			ReDim preserve stat_disq_four_type_info(memb_count)
			ReDim preserve stat_disq_four_begin_date(memb_count)
			ReDim preserve stat_disq_four_end_date(memb_count)
			ReDim preserve stat_disq_four_cure_reason_code(memb_count)
			ReDim preserve stat_disq_four_cure_reason_info(memb_count)
			ReDim preserve stat_disq_four_fraud_determination_date(memb_count)
			ReDim preserve stat_disq_four_county_of_fraud(memb_count)
			ReDim preserve stat_disq_four_state_of_fraud(memb_count)
			ReDim preserve stat_disq_four_SNAP_trafficking_yn(memb_count)
			ReDim preserve stat_disq_four_SNAP_offense_code(memb_count)
			ReDim preserve stat_disq_four_SNAP_offense_info(memb_count)
			ReDim preserve stat_disq_four_source(memb_count)
			ReDim preserve stat_disq_four_active(memb_count)

			ReDim preserve stat_disq_five_exists(memb_count)
			ReDim preserve stat_disq_five_program(memb_count)
			ReDim preserve stat_disq_five_type_code(memb_count)
			ReDim preserve stat_disq_five_type_info(memb_count)
			ReDim preserve stat_disq_five_begin_date(memb_count)
			ReDim preserve stat_disq_five_end_date(memb_count)
			ReDim preserve stat_disq_five_cure_reason_code(memb_count)
			ReDim preserve stat_disq_five_cure_reason_info(memb_count)
			ReDim preserve stat_disq_five_fraud_determination_date(memb_count)
			ReDim preserve stat_disq_five_county_of_fraud(memb_count)
			ReDim preserve stat_disq_five_state_of_fraud(memb_count)
			ReDim preserve stat_disq_five_SNAP_trafficking_yn(memb_count)
			ReDim preserve stat_disq_five_SNAP_offense_code(memb_count)
			ReDim preserve stat_disq_five_SNAP_offense_info(memb_count)
			ReDim preserve stat_disq_five_source(memb_count)
			ReDim preserve stat_disq_five_active(memb_count)

			EMReadScreen stat_memb_ref_numb(memb_count), 2, 4, 33
			EMReadScreen stat_memb_last_name(memb_count), 25, 6, 30
			EMReadScreen stat_memb_first_name(memb_count), 12, 6, 63
			EMReadScreen stat_memb_middle_initial(memb_count), 1, 6, 79

			stat_memb_first_name(memb_count) = replace(stat_memb_first_name(memb_count), "_", "")
			stat_memb_last_name(memb_count) = replace(stat_memb_last_name(memb_count), "_", "")
			stat_memb_middle_initial(memb_count) = replace(stat_memb_middle_initial(memb_count), "_", "")

			stat_memb_full_name(memb_count) = stat_memb_first_name(memb_count) & " " & stat_memb_middle_initial(memb_count) & ". " & stat_memb_last_name(memb_count)
			stat_memb_full_name_no_initial(memb_count) = stat_memb_first_name(memb_count) & " " & stat_memb_last_name(memb_count)
			stat_memb_full_name_last_name_first(memb_count) = stat_memb_last_name(memb_count) & ", " & stat_memb_first_name(memb_count) & " " & stat_memb_middle_initial(memb_count)
			stat_memb_full_name_last_name_first_no_mi(memb_count) = stat_memb_last_name(memb_count) & ", " & stat_memb_first_name(memb_count)

			EMReadScreen stat_memb_age(memb_count), 3, 8, 76
			EMReadScreen stat_memb_id_verif_code(memb_count), 2, 9, 68
			EMReadScreen stat_memb_rel_to_applct_code(memb_count), 2, 10, 42

			stat_memb_age(memb_count) = trim(stat_memb_age(memb_count))
			If stat_memb_age(memb_count) = "" Then stat_memb_age(memb_count) = 0
			stat_memb_age(memb_count) = stat_memb_age(memb_count)*1

			If stat_memb_id_verif_code(memb_count) = "BC" Then stat_memb_id_verif_info(memb_count) = "Birth Certificate"
			If stat_memb_id_verif_code(memb_count) = "RE" Then stat_memb_id_verif_info(memb_count) = "Religious Record"
			If stat_memb_id_verif_code(memb_count) = "DL" Then stat_memb_id_verif_info(memb_count) = "Drivers License/State ID"
			If stat_memb_id_verif_code(memb_count) = "DV" Then stat_memb_id_verif_info(memb_count) = "Divorce Decree"
			If stat_memb_id_verif_code(memb_count) = "AL" Then stat_memb_id_verif_info(memb_count) = "Alien ID Card"
			If stat_memb_id_verif_code(memb_count) = "AD" Then stat_memb_id_verif_info(memb_count) = "Arrival/Departure Document - I94"
			If stat_memb_id_verif_code(memb_count) = "DR" Then stat_memb_id_verif_info(memb_count) = "Doctor Statement"
			If stat_memb_id_verif_code(memb_count) = "PV" Then stat_memb_id_verif_info(memb_count) = "Passport/Visa"
			If stat_memb_id_verif_code(memb_count) = "OT" Then stat_memb_id_verif_info(memb_count) = "Other Document"
			If stat_memb_id_verif_code(memb_count) = "NO" Then stat_memb_id_verif_info(memb_count) = "No Verif Provided"

			If stat_memb_rel_to_applct_code(memb_count) = "01" Then stat_memb_rel_to_applct_info(memb_count) = "Applicant"
			If stat_memb_rel_to_applct_code(memb_count) = "02" Then stat_memb_rel_to_applct_info(memb_count) = "Spouse"
			If stat_memb_rel_to_applct_code(memb_count) = "03" Then stat_memb_rel_to_applct_info(memb_count) = "Child"
			If stat_memb_rel_to_applct_code(memb_count) = "04" Then stat_memb_rel_to_applct_info(memb_count) = "Parent"
			If stat_memb_rel_to_applct_code(memb_count) = "05" Then stat_memb_rel_to_applct_info(memb_count) = "Sibling"
			If stat_memb_rel_to_applct_code(memb_count) = "06" Then stat_memb_rel_to_applct_info(memb_count) = "Step Sibling"
			If stat_memb_rel_to_applct_code(memb_count) = "08" Then stat_memb_rel_to_applct_info(memb_count) = "Step Child"
			If stat_memb_rel_to_applct_code(memb_count) = "09" Then stat_memb_rel_to_applct_info(memb_count) = "Step Parent"
			If stat_memb_rel_to_applct_code(memb_count) = "10" Then stat_memb_rel_to_applct_info(memb_count) = "Aunt"
			If stat_memb_rel_to_applct_code(memb_count) = "11" Then stat_memb_rel_to_applct_info(memb_count) = "Uncle"
			If stat_memb_rel_to_applct_code(memb_count) = "12" Then stat_memb_rel_to_applct_info(memb_count) = "Niece"
			If stat_memb_rel_to_applct_code(memb_count) = "13" Then stat_memb_rel_to_applct_info(memb_count) = "Nephew"
			If stat_memb_rel_to_applct_code(memb_count) = "14" Then stat_memb_rel_to_applct_info(memb_count) = "Cousin"
			If stat_memb_rel_to_applct_code(memb_count) = "15" Then stat_memb_rel_to_applct_info(memb_count) = "Grandparent"
			If stat_memb_rel_to_applct_code(memb_count) = "16" Then stat_memb_rel_to_applct_info(memb_count) = "Grandchild"
			If stat_memb_rel_to_applct_code(memb_count) = "17" Then stat_memb_rel_to_applct_info(memb_count) = "Other Relative"
			If stat_memb_rel_to_applct_code(memb_count) = "18" Then stat_memb_rel_to_applct_info(memb_count) = "Legal Guardian"
			If stat_memb_rel_to_applct_code(memb_count) = "24" Then stat_memb_rel_to_applct_info(memb_count) = "Not Related"
			If stat_memb_rel_to_applct_code(memb_count) = "25" Then stat_memb_rel_to_applct_info(memb_count) = "Live-In Attendant"
			If stat_memb_rel_to_applct_code(memb_count) = "27" Then stat_memb_rel_to_applct_info(memb_count) = "Unknown"

			transmit
			EMReadScreen next_ref_numb, 2, 4, 33
		Loop until next_ref_numb = stat_memb_ref_numb(memb_count)

		Call navigate_to_MAXIS_screen("STAT", "MEMI")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			Call write_value_and_transmit(stat_memb_ref_numb(each_memb), 20, 76)

			EMReadScreen stat_memi_citizenship_yn(each_memb), 1, 11, 49
			EMReadScreen stat_memi_citizenship_verif_code(each_memb), 2, 11, 78

			If stat_memi_citizenship_verif_code(each_memb) = "BC" Then stat_memi_citizenship_verif_info(each_memb) = "Birth Certificate"
			If stat_memi_citizenship_verif_code(each_memb) = "RE" Then stat_memi_citizenship_verif_info(each_memb) = "Religious Record"
			If stat_memi_citizenship_verif_code(each_memb) = "NP" Then stat_memi_citizenship_verif_info(each_memb) = "Naturalization Papers"
			If stat_memi_citizenship_verif_code(each_memb) = "IM" Then stat_memi_citizenship_verif_info(each_memb) = "Immigration Document"
			If stat_memi_citizenship_verif_code(each_memb) = "PV" Then stat_memi_citizenship_verif_info(each_memb) = "Passport/Visa"
			If stat_memi_citizenship_verif_code(each_memb) = "OT" Then stat_memi_citizenship_verif_info(each_memb) = "Other Document"
			If stat_memi_citizenship_verif_code(each_memb) = "NO" Then stat_memi_citizenship_verif_info(each_memb) = "No Verif Provided"
		Next

		Call navigate_to_MAXIS_screen("STAT", "JOBS")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_jobs_one_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_jobs_one_exists(each_memb) = False

			If stat_jobs_one_exists(each_memb) = True Then
				EMReadScreen stat_jobs_one_inc_type(each_memb), 1, 5, 34
				EMReadScreen stat_jobs_one_sub_inc_type(each_memb), 2, 5, 74
				EMReadScreen stat_jobs_one_verif_code(each_memb), 1, 6, 34
				EMReadScreen stat_jobs_one_employer_name(each_memb), 30, 7, 42
				EMReadScreen stat_jobs_one_inc_start_date(each_memb), 8, 9, 35
				EMReadScreen stat_jobs_one_inc_end_date(each_memb), 8, 9, 49
				EMReadScreen stat_jobs_one_main_pay_freq(each_memb), 1, 18, 35

				If stat_jobs_one_verif_code(each_memb) = "1" Then stat_jobs_one_verif_info(each_memb) = "Pay Stubs/Tip Report"
				If stat_jobs_one_verif_code(each_memb) = "2" Then stat_jobs_one_verif_info(each_memb) = "Employer Statement"
				If stat_jobs_one_verif_code(each_memb) = "3" Then stat_jobs_one_verif_info(each_memb) = "Collateral Statement"
				If stat_jobs_one_verif_code(each_memb) = "4" Then stat_jobs_one_verif_info(each_memb) = "Other Document"
				If stat_jobs_one_verif_code(each_memb) = "5" Then stat_jobs_one_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_jobs_one_verif_code(each_memb) = "N" Then stat_jobs_one_verif_info(each_memb) = "No Verif Provided"

				stat_jobs_one_employer_name(each_memb) = replace(stat_jobs_one_employer_name(each_memb), "_", "")
				stat_jobs_one_inc_start_date(each_memb) = replace(stat_jobs_one_inc_start_date(each_memb), " ", "/")
				stat_jobs_one_inc_end_date(each_memb) = replace(stat_jobs_one_inc_end_date(each_memb), " ", "/")
				If stat_jobs_one_inc_end_date(each_memb) = "__/__/__" Then stat_jobs_one_inc_end_date(each_memb) = ""

				stat_jobs_one_job_ended(each_memb) = False
				stat_jobs_one_job_counted(each_memb) = True
				If IsDate(stat_jobs_one_inc_end_date(each_memb)) = True Then
					If DateDiff("m", stat_jobs_one_inc_end_date(each_memb), current_month) > 0 Then stat_jobs_one_job_ended(each_memb) = True
				End If
				If stat_jobs_one_job_ended(each_memb) = True Then stat_jobs_one_job_counted(each_memb) = False
				If stat_jobs_one_inc_type(each_memb) = "G" Then stat_jobs_one_job_counted(each_memb) = False
				If stat_jobs_one_inc_type(each_memb) = "F" Then stat_jobs_one_job_counted(each_memb) = False
				If stat_jobs_one_inc_type(each_memb) = "S" Then stat_jobs_one_job_counted(each_memb) = False

				If stat_jobs_one_sub_inc_type(each_memb) = "04" Then stat_jobs_one_job_counted(each_memb) = False

				If stat_jobs_one_main_pay_freq(each_memb) = "1" Then stat_jobs_one_main_pay_freq(each_memb) = "Monthly"
				If stat_jobs_one_main_pay_freq(each_memb) = "2" Then stat_jobs_one_main_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_one_main_pay_freq(each_memb) = "3" Then stat_jobs_one_main_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_one_main_pay_freq(each_memb) = "4" Then stat_jobs_one_main_pay_freq(each_memb) = "Weekly"
				If stat_jobs_one_main_pay_freq(each_memb) = "5" Then stat_jobs_one_main_pay_freq(each_memb) = "Other"

				Call write_value_and_transmit("X", 19, 38)
				EMReadScreen stat_jobs_one_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_jobs_one_snap_pic_ave_hrs_per_pay(each_memb), 7, 16, 50
				EMReadScreen stat_jobs_one_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 54
				EMReadScreen stat_jobs_one_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 54

				If stat_jobs_one_snap_pic_pay_freq(each_memb) = "1" Then stat_jobs_one_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_one_snap_pic_pay_freq(each_memb) = "2" Then stat_jobs_one_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_one_snap_pic_pay_freq(each_memb) = "3" Then stat_jobs_one_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_one_snap_pic_pay_freq(each_memb) = "4" Then stat_jobs_one_snap_pic_pay_freq(each_memb) = "Weekly"
				If stat_jobs_one_snap_pic_pay_freq(each_memb) = "5" Then stat_jobs_one_snap_pic_pay_freq(each_memb) = "Other"

				stat_jobs_one_snap_pic_ave_hrs_per_pay(each_memb) = trim(stat_jobs_one_snap_pic_ave_hrs_per_pay(each_memb))
				stat_jobs_one_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_one_snap_pic_ave_inc_per_pay(each_memb))
				stat_jobs_one_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_one_snap_pic_prosp_monthly_inc(each_memb))
				PF3

				Call write_value_and_transmit("X", 19, 71)
				EMReadScreen stat_jobs_one_grh_pic_pay_freq(each_memb), 1, 3, 63
				EMReadScreen stat_jobs_one_grh_pic_ave_inc_per_pay(each_memb), 10, 16, 65
				EMReadScreen stat_jobs_one_grh_pic_prosp_monthly_inc(each_memb), 10, 17, 65

				If stat_jobs_one_grh_pic_pay_freq(each_memb) = "1" Then stat_jobs_one_grh_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_one_grh_pic_pay_freq(each_memb) = "2" Then stat_jobs_one_grh_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_one_grh_pic_pay_freq(each_memb) = "3" Then stat_jobs_one_grh_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_one_grh_pic_pay_freq(each_memb) = "4" Then stat_jobs_one_grh_pic_pay_freq(each_memb) = "Weekly"

				stat_jobs_one_grh_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_one_grh_pic_ave_inc_per_pay(each_memb))
				stat_jobs_one_grh_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_one_grh_pic_prosp_monthly_inc(each_memb))
				PF3
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "02", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_jobs_two_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_jobs_two_exists(each_memb) = False

			If stat_jobs_two_exists(each_memb) = True Then
				EMReadScreen stat_jobs_two_verif_code(each_memb), 1, 6, 34
				EMReadScreen stat_jobs_two_employer_name(each_memb), 30, 7, 42
				EMReadScreen stat_jobs_two_inc_start_date(each_memb), 8, 9, 35
				EMReadScreen stat_jobs_two_inc_end_date(each_memb), 8, 9, 49
				EMReadScreen stat_jobs_two_main_pay_freq(each_memb), 1, 18, 35

				If stat_jobs_two_verif_code(each_memb) = "1" Then stat_jobs_two_verif_info(each_memb) = "Pay Stubs/Tip Report"
				If stat_jobs_two_verif_code(each_memb) = "2" Then stat_jobs_two_verif_info(each_memb) = "Employer Statement"
				If stat_jobs_two_verif_code(each_memb) = "3" Then stat_jobs_two_verif_info(each_memb) = "Collateral Statement"
				If stat_jobs_two_verif_code(each_memb) = "4" Then stat_jobs_two_verif_info(each_memb) = "Other Document"
				If stat_jobs_two_verif_code(each_memb) = "5" Then stat_jobs_two_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_jobs_two_verif_code(each_memb) = "N" Then stat_jobs_two_verif_info(each_memb) = "No Verif Provided"

				stat_jobs_two_employer_name(each_memb) = replace(stat_jobs_two_employer_name(each_memb), "_", "")
				stat_jobs_two_inc_start_date(each_memb) = replace(stat_jobs_two_inc_start_date(each_memb), " ", "/")
				stat_jobs_two_inc_end_date(each_memb) = replace(stat_jobs_two_inc_end_date(each_memb), " ", "/")
				If stat_jobs_two_inc_end_date(each_memb) = "__/__/__" Then stat_jobs_two_inc_end_date(each_memb) = ""

				stat_jobs_two_job_ended(each_memb) = False
				stat_jobs_two_job_counted(each_memb) = True
				If IsDate(stat_jobs_two_inc_end_date(each_memb)) = True Then
					If DateDiff("m", stat_jobs_two_inc_end_date(each_memb), current_month) > 0 Then stat_jobs_two_job_ended(each_memb) = True
				End If
				If stat_jobs_two_job_ended(each_memb) = True Then stat_jobs_two_job_counted(each_memb) = False
				If stat_jobs_two_inc_type(each_memb) = "G" Then stat_jobs_two_job_counted(each_memb) = False
				If stat_jobs_two_inc_type(each_memb) = "F" Then stat_jobs_two_job_counted(each_memb) = False
				If stat_jobs_two_inc_type(each_memb) = "S" Then stat_jobs_two_job_counted(each_memb) = False

				If stat_jobs_two_sub_inc_type(each_memb) = "04" Then stat_jobs_two_job_counted(each_memb) = False

				If stat_jobs_two_main_pay_freq(each_memb) = "1" Then stat_jobs_two_main_pay_freq(each_memb) = "Monthly"
				If stat_jobs_two_main_pay_freq(each_memb) = "2" Then stat_jobs_two_main_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_two_main_pay_freq(each_memb) = "3" Then stat_jobs_two_main_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_two_main_pay_freq(each_memb) = "4" Then stat_jobs_two_main_pay_freq(each_memb) = "Weekly"
				If stat_jobs_two_main_pay_freq(each_memb) = "5" Then stat_jobs_two_main_pay_freq(each_memb) = "Other"

				Call write_value_and_transmit("X", 19, 38)
				EMReadScreen stat_jobs_two_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_jobs_two_snap_pic_ave_hrs_per_pay(each_memb), 7, 16, 50
				EMReadScreen stat_jobs_two_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 54
				EMReadScreen stat_jobs_two_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 54

				If stat_jobs_two_snap_pic_pay_freq(each_memb) = "1" Then stat_jobs_two_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_two_snap_pic_pay_freq(each_memb) = "2" Then stat_jobs_two_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_two_snap_pic_pay_freq(each_memb) = "3" Then stat_jobs_two_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_two_snap_pic_pay_freq(each_memb) = "4" Then stat_jobs_two_snap_pic_pay_freq(each_memb) = "Weekly"
				If stat_jobs_two_snap_pic_pay_freq(each_memb) = "5" Then stat_jobs_two_snap_pic_pay_freq(each_memb) = "Other"

				stat_jobs_two_snap_pic_ave_hrs_per_pay(each_memb) = trim(stat_jobs_two_snap_pic_ave_hrs_per_pay(each_memb))
				stat_jobs_two_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_two_snap_pic_ave_inc_per_pay(each_memb))
				stat_jobs_two_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_two_snap_pic_prosp_monthly_inc(each_memb))
				PF3

				Call write_value_and_transmit("X", 19, 71)
				EMReadScreen stat_jobs_two_grh_pic_pay_freq(each_memb), 1, 3, 63
				EMReadScreen stat_jobs_two_grh_pic_ave_inc_per_pay(each_memb), 10, 16, 65
				EMReadScreen stat_jobs_two_grh_pic_prosp_monthly_inc(each_memb), 10, 17, 65

				If stat_jobs_two_grh_pic_pay_freq(each_memb) = "1" Then stat_jobs_two_grh_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_two_grh_pic_pay_freq(each_memb) = "2" Then stat_jobs_two_grh_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_two_grh_pic_pay_freq(each_memb) = "3" Then stat_jobs_two_grh_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_two_grh_pic_pay_freq(each_memb) = "4" Then stat_jobs_two_grh_pic_pay_freq(each_memb) = "Weekly"

				stat_jobs_two_grh_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_two_grh_pic_ave_inc_per_pay(each_memb))
				stat_jobs_two_grh_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_two_grh_pic_prosp_monthly_inc(each_memb))
				PF3
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "03", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_jobs_three_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_jobs_three_exists(each_memb) = False

			If stat_jobs_three_exists(each_memb) = True Then
				EMReadScreen stat_jobs_three_verif_code(each_memb), 1, 6, 34
				EMReadScreen stat_jobs_three_employer_name(each_memb), 30, 7, 42
				EMReadScreen stat_jobs_three_inc_start_date(each_memb), 8, 9, 35
				EMReadScreen stat_jobs_three_inc_end_date(each_memb), 8, 9, 49
				EMReadScreen stat_jobs_three_main_pay_freq(each_memb), 1, 18, 35

				If stat_jobs_three_verif_code(each_memb) = "1" Then stat_jobs_three_verif_info(each_memb) = "Pay Stubs/Tip Report"
				If stat_jobs_three_verif_code(each_memb) = "2" Then stat_jobs_three_verif_info(each_memb) = "Employer Statement"
				If stat_jobs_three_verif_code(each_memb) = "3" Then stat_jobs_three_verif_info(each_memb) = "Collateral Statement"
				If stat_jobs_three_verif_code(each_memb) = "4" Then stat_jobs_three_verif_info(each_memb) = "Other Document"
				If stat_jobs_three_verif_code(each_memb) = "5" Then stat_jobs_three_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_jobs_three_verif_code(each_memb) = "N" Then stat_jobs_three_verif_info(each_memb) = "No Verif Provided"

				stat_jobs_three_employer_name(each_memb) = replace(stat_jobs_three_employer_name(each_memb), "_", "")
				stat_jobs_three_inc_start_date(each_memb) = replace(stat_jobs_three_inc_start_date(each_memb), " ", "/")
				stat_jobs_three_inc_end_date(each_memb) = replace(stat_jobs_three_inc_end_date(each_memb), " ", "/")
				If stat_jobs_three_inc_end_date(each_memb) = "__/__/__" Then stat_jobs_three_inc_end_date(each_memb) = ""

				stat_jobs_three_job_ended(each_memb) = False
				stat_jobs_three_job_counted(each_memb) = True
				If IsDate(stat_jobs_three_inc_end_date(each_memb)) = True Then
					If DateDiff("m", stat_jobs_three_inc_end_date(each_memb), current_month) > 0 Then stat_jobs_three_job_ended(each_memb) = True
				End If
				If stat_jobs_three_job_ended(each_memb) = True Then stat_jobs_three_job_counted(each_memb) = False
				If stat_jobs_three_inc_type(each_memb) = "G" Then stat_jobs_three_job_counted(each_memb) = False
				If stat_jobs_three_inc_type(each_memb) = "F" Then stat_jobs_three_job_counted(each_memb) = False
				If stat_jobs_three_inc_type(each_memb) = "S" Then stat_jobs_three_job_counted(each_memb) = False

				If stat_jobs_three_sub_inc_type(each_memb) = "04" Then stat_jobs_three_job_counted(each_memb) = False

				If stat_jobs_three_main_pay_freq(each_memb) = "1" Then stat_jobs_three_main_pay_freq(each_memb) = "Monthly"
				If stat_jobs_three_main_pay_freq(each_memb) = "2" Then stat_jobs_three_main_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_three_main_pay_freq(each_memb) = "3" Then stat_jobs_three_main_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_three_main_pay_freq(each_memb) = "4" Then stat_jobs_three_main_pay_freq(each_memb) = "Weekly"
				If stat_jobs_three_main_pay_freq(each_memb) = "5" Then stat_jobs_three_main_pay_freq(each_memb) = "Other"

				Call write_value_and_transmit("X", 19, 38)
				EMReadScreen stat_jobs_three_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_jobs_three_snap_pic_ave_hrs_per_pay(each_memb), 7, 16, 50
				EMReadScreen stat_jobs_three_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 54
				EMReadScreen stat_jobs_three_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 54

				If stat_jobs_three_snap_pic_pay_freq(each_memb) = "1" Then stat_jobs_three_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_three_snap_pic_pay_freq(each_memb) = "2" Then stat_jobs_three_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_three_snap_pic_pay_freq(each_memb) = "3" Then stat_jobs_three_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_three_snap_pic_pay_freq(each_memb) = "4" Then stat_jobs_three_snap_pic_pay_freq(each_memb) = "Weekly"
				If stat_jobs_three_snap_pic_pay_freq(each_memb) = "5" Then stat_jobs_three_snap_pic_pay_freq(each_memb) = "Other"

				stat_jobs_three_snap_pic_ave_hrs_per_pay(each_memb) = trim(stat_jobs_three_snap_pic_ave_hrs_per_pay(each_memb))
				stat_jobs_three_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_three_snap_pic_ave_inc_per_pay(each_memb))
				stat_jobs_three_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_three_snap_pic_prosp_monthly_inc(each_memb))
				PF3

				Call write_value_and_transmit("X", 19, 71)
				EMReadScreen stat_jobs_three_grh_pic_pay_freq(each_memb), 1, 3, 63
				EMReadScreen stat_jobs_three_grh_pic_ave_inc_per_pay(each_memb), 10, 16, 65
				EMReadScreen stat_jobs_three_grh_pic_prosp_monthly_inc(each_memb), 10, 17, 65

				If stat_jobs_three_grh_pic_pay_freq(each_memb) = "1" Then stat_jobs_three_grh_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_three_grh_pic_pay_freq(each_memb) = "2" Then stat_jobs_three_grh_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_three_grh_pic_pay_freq(each_memb) = "3" Then stat_jobs_three_grh_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_three_grh_pic_pay_freq(each_memb) = "4" Then stat_jobs_three_grh_pic_pay_freq(each_memb) = "Weekly"

				stat_jobs_three_grh_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_three_grh_pic_ave_inc_per_pay(each_memb))
				stat_jobs_three_grh_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_three_grh_pic_prosp_monthly_inc(each_memb))
				PF3
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "04", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_jobs_four_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_jobs_four_exists(each_memb) = False

			If stat_jobs_four_exists(each_memb) = True Then
				EMReadScreen stat_jobs_four_verif_code(each_memb), 1, 6, 34
				EMReadScreen stat_jobs_four_employer_name(each_memb), 30, 7, 42
				EMReadScreen stat_jobs_four_inc_start_date(each_memb), 8, 9, 35
				EMReadScreen stat_jobs_four_inc_end_date(each_memb), 8, 9, 49
				EMReadScreen stat_jobs_four_main_pay_freq(each_memb), 1, 18, 35

				If stat_jobs_four_verif_code(each_memb) = "1" Then stat_jobs_four_verif_info(each_memb) = "Pay Stubs/Tip Report"
				If stat_jobs_four_verif_code(each_memb) = "2" Then stat_jobs_four_verif_info(each_memb) = "Employer Statement"
				If stat_jobs_four_verif_code(each_memb) = "3" Then stat_jobs_four_verif_info(each_memb) = "Collateral Statement"
				If stat_jobs_four_verif_code(each_memb) = "4" Then stat_jobs_four_verif_info(each_memb) = "Other Document"
				If stat_jobs_four_verif_code(each_memb) = "5" Then stat_jobs_four_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_jobs_four_verif_code(each_memb) = "N" Then stat_jobs_four_verif_info(each_memb) = "No Verif Provided"

				stat_jobs_four_employer_name(each_memb) = replace(stat_jobs_four_employer_name(each_memb), "_", "")
				stat_jobs_four_inc_start_date(each_memb) = replace(stat_jobs_four_inc_start_date(each_memb), " ", "/")
				stat_jobs_four_inc_end_date(each_memb) = replace(stat_jobs_four_inc_end_date(each_memb), " ", "/")
				If stat_jobs_four_inc_end_date(each_memb) = "__/__/__" Then stat_jobs_four_inc_end_date(each_memb) = ""

				stat_jobs_four_job_ended(each_memb) = False
				stat_jobs_four_job_counted(each_memb) = True
				If IsDate(stat_jobs_four_inc_end_date(each_memb)) = True Then
					If DateDiff("m", stat_jobs_four_inc_end_date(each_memb), current_month) > 0 Then stat_jobs_four_job_ended(each_memb) = True
				End If
				If stat_jobs_four_job_ended(each_memb) = True Then stat_jobs_four_job_counted(each_memb) = False
				If stat_jobs_four_inc_type(each_memb) = "G" Then stat_jobs_four_job_counted(each_memb) = False
				If stat_jobs_four_inc_type(each_memb) = "F" Then stat_jobs_four_job_counted(each_memb) = False
				If stat_jobs_four_inc_type(each_memb) = "S" Then stat_jobs_four_job_counted(each_memb) = False

				If stat_jobs_four_sub_inc_type(each_memb) = "04" Then stat_jobs_four_job_counted(each_memb) = False

				If stat_jobs_four_main_pay_freq(each_memb) = "1" Then stat_jobs_four_main_pay_freq(each_memb) = "Monthly"
				If stat_jobs_four_main_pay_freq(each_memb) = "2" Then stat_jobs_four_main_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_four_main_pay_freq(each_memb) = "3" Then stat_jobs_four_main_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_four_main_pay_freq(each_memb) = "4" Then stat_jobs_four_main_pay_freq(each_memb) = "Weekly"
				If stat_jobs_four_main_pay_freq(each_memb) = "5" Then stat_jobs_four_main_pay_freq(each_memb) = "Other"

				Call write_value_and_transmit("X", 19, 38)
				EMReadScreen stat_jobs_four_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_jobs_four_snap_pic_ave_hrs_per_pay(each_memb), 7, 16, 50
				EMReadScreen stat_jobs_four_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 54
				EMReadScreen stat_jobs_four_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 54

				If stat_jobs_four_snap_pic_pay_freq(each_memb) = "1" Then stat_jobs_four_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_four_snap_pic_pay_freq(each_memb) = "2" Then stat_jobs_four_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_four_snap_pic_pay_freq(each_memb) = "3" Then stat_jobs_four_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_four_snap_pic_pay_freq(each_memb) = "4" Then stat_jobs_four_snap_pic_pay_freq(each_memb) = "Weekly"
				If stat_jobs_four_snap_pic_pay_freq(each_memb) = "5" Then stat_jobs_four_snap_pic_pay_freq(each_memb) = "Other"

				stat_jobs_four_snap_pic_ave_hrs_per_pay(each_memb) = trim(stat_jobs_four_snap_pic_ave_hrs_per_pay(each_memb))
				stat_jobs_four_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_four_snap_pic_ave_inc_per_pay(each_memb))
				stat_jobs_four_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_four_snap_pic_prosp_monthly_inc(each_memb))
				PF3

				Call write_value_and_transmit("X", 19, 71)
				EMReadScreen stat_jobs_four_grh_pic_pay_freq(each_memb), 1, 3, 63
				EMReadScreen stat_jobs_four_grh_pic_ave_inc_per_pay(each_memb), 10, 16, 65
				EMReadScreen stat_jobs_four_grh_pic_prosp_monthly_inc(each_memb), 10, 17, 65

				If stat_jobs_four_grh_pic_pay_freq(each_memb) = "1" Then stat_jobs_four_grh_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_four_grh_pic_pay_freq(each_memb) = "2" Then stat_jobs_four_grh_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_four_grh_pic_pay_freq(each_memb) = "3" Then stat_jobs_four_grh_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_four_grh_pic_pay_freq(each_memb) = "4" Then stat_jobs_four_grh_pic_pay_freq(each_memb) = "Weekly"

				stat_jobs_four_grh_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_four_grh_pic_ave_inc_per_pay(each_memb))
				stat_jobs_four_grh_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_four_grh_pic_prosp_monthly_inc(each_memb))
				PF3
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "05", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_jobs_five_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_jobs_five_exists(each_memb) = False

			If stat_jobs_five_exists(each_memb) = True Then
				EMReadScreen stat_jobs_five_verif_code(each_memb), 1, 6, 34
				EMReadScreen stat_jobs_five_employer_name(each_memb), 30, 7, 42
				EMReadScreen stat_jobs_five_inc_start_date(each_memb), 8, 9, 35
				EMReadScreen stat_jobs_five_inc_end_date(each_memb), 8, 9, 49
				EMReadScreen stat_jobs_five_main_pay_freq(each_memb), 1, 18, 35

				If stat_jobs_five_verif_code(each_memb) = "1" Then stat_jobs_five_verif_info(each_memb) = "Pay Stubs/Tip Report"
				If stat_jobs_five_verif_code(each_memb) = "2" Then stat_jobs_five_verif_info(each_memb) = "Employer Statement"
				If stat_jobs_five_verif_code(each_memb) = "3" Then stat_jobs_five_verif_info(each_memb) = "Collateral Statement"
				If stat_jobs_five_verif_code(each_memb) = "4" Then stat_jobs_five_verif_info(each_memb) = "Other Document"
				If stat_jobs_five_verif_code(each_memb) = "5" Then stat_jobs_five_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_jobs_five_verif_code(each_memb) = "N" Then stat_jobs_five_verif_info(each_memb) = "No Verif Provided"

				stat_jobs_five_employer_name(each_memb) = replace(stat_jobs_five_employer_name(each_memb), "_", "")
				stat_jobs_five_inc_start_date(each_memb) = replace(stat_jobs_five_inc_start_date(each_memb), " ", "/")
				stat_jobs_five_inc_end_date(each_memb) = replace(stat_jobs_five_inc_end_date(each_memb), " ", "/")
				If stat_jobs_five_inc_end_date(each_memb) = "__/__/__" Then stat_jobs_five_inc_end_date(each_memb) = ""

				stat_jobs_five_job_ended(each_memb) = False
				stat_jobs_five_job_counted(each_memb) = True
				If IsDate(stat_jobs_five_inc_end_date(each_memb)) = True Then
					If DateDiff("m", stat_jobs_five_inc_end_date(each_memb), current_month) > 0 Then stat_jobs_five_job_ended(each_memb) = True
				End If
				If stat_jobs_five_job_ended(each_memb) = True Then stat_jobs_five_job_counted(each_memb) = False
				If stat_jobs_five_inc_type(each_memb) = "G" Then stat_jobs_five_job_counted(each_memb) = False
				If stat_jobs_five_inc_type(each_memb) = "F" Then stat_jobs_five_job_counted(each_memb) = False
				If stat_jobs_five_inc_type(each_memb) = "S" Then stat_jobs_five_job_counted(each_memb) = False

				If stat_jobs_five_sub_inc_type(each_memb) = "04" Then stat_jobs_five_job_counted(each_memb) = False

				If stat_jobs_five_main_pay_freq(each_memb) = "1" Then stat_jobs_five_main_pay_freq(each_memb) = "Monthly"
				If stat_jobs_five_main_pay_freq(each_memb) = "2" Then stat_jobs_five_main_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_five_main_pay_freq(each_memb) = "3" Then stat_jobs_five_main_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_five_main_pay_freq(each_memb) = "4" Then stat_jobs_five_main_pay_freq(each_memb) = "Weekly"
				If stat_jobs_five_main_pay_freq(each_memb) = "5" Then stat_jobs_five_main_pay_freq(each_memb) = "Other"

				Call write_value_and_transmit("X", 19, 38)
				EMReadScreen stat_jobs_five_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_jobs_five_snap_pic_ave_hrs_per_pay(each_memb), 7, 16, 50
				EMReadScreen stat_jobs_five_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 54
				EMReadScreen stat_jobs_five_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 54

				If stat_jobs_five_snap_pic_pay_freq(each_memb) = "1" Then stat_jobs_five_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_five_snap_pic_pay_freq(each_memb) = "2" Then stat_jobs_five_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_five_snap_pic_pay_freq(each_memb) = "3" Then stat_jobs_five_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_five_snap_pic_pay_freq(each_memb) = "4" Then stat_jobs_five_snap_pic_pay_freq(each_memb) = "Weekly"
				If stat_jobs_five_snap_pic_pay_freq(each_memb) = "5" Then stat_jobs_five_snap_pic_pay_freq(each_memb) = "Other"

				stat_jobs_five_snap_pic_ave_hrs_per_pay(each_memb) = trim(stat_jobs_five_snap_pic_ave_hrs_per_pay(each_memb))
				stat_jobs_five_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_five_snap_pic_ave_inc_per_pay(each_memb))
				stat_jobs_five_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_five_snap_pic_prosp_monthly_inc(each_memb))
				PF3

				Call write_value_and_transmit("X", 19, 71)
				EMReadScreen stat_jobs_five_grh_pic_pay_freq(each_memb), 1, 3, 63
				EMReadScreen stat_jobs_five_grh_pic_ave_inc_per_pay(each_memb), 10, 16, 65
				EMReadScreen stat_jobs_five_grh_pic_prosp_monthly_inc(each_memb), 10, 17, 65

				If stat_jobs_five_grh_pic_pay_freq(each_memb) = "1" Then stat_jobs_five_grh_pic_pay_freq(each_memb) = "Monthly"
				If stat_jobs_five_grh_pic_pay_freq(each_memb) = "2" Then stat_jobs_five_grh_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_jobs_five_grh_pic_pay_freq(each_memb) = "3" Then stat_jobs_five_grh_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_jobs_five_grh_pic_pay_freq(each_memb) = "4" Then stat_jobs_five_grh_pic_pay_freq(each_memb) = "Weekly"

				stat_jobs_five_grh_pic_ave_inc_per_pay(each_memb) = trim(stat_jobs_five_grh_pic_ave_inc_per_pay(each_memb))
				stat_jobs_five_grh_pic_prosp_monthly_inc(each_memb) = trim(stat_jobs_five_grh_pic_prosp_monthly_inc(each_memb))
				PF3
			End If
		Next

		call navigate_to_MAXIS_screen("STAT", "BUSI")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_busi_one_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_busi_one_exists(each_memb) = False

			If stat_busi_one_exists(each_memb) = True Then
				stat_busi_one_counted(each_memb) = True
				EMReadScreen stat_busi_one_type(each_memb), 2, 5, 37
				EMReadScreen stat_busi_one_inc_start_date(each_memb), 8, 5, 55
				EMReadScreen stat_busi_one_inc_end_date(each_memb), 8, 5, 72
				EMReadScreen stat_busi_one_method(each_memb), 2, 16, 53
				EMReadScreen stat_busi_one_method_date(each_memb), 8, 16, 63

				If stat_busi_one_type(each_memb) = "01" Then stat_busi_one_type_info(each_memb) = "Farming"
				If stat_busi_one_type(each_memb) = "02" Then stat_busi_one_type_info(each_memb) = "Real Estate"
				If stat_busi_one_type(each_memb) = "03" Then stat_busi_one_type_info(each_memb) = "Home Product Sales"
				If stat_busi_one_type(each_memb) = "04" Then stat_busi_one_type_info(each_memb) = "Other Sales"
				If stat_busi_one_type(each_memb) = "05" Then stat_busi_one_type_info(each_memb) = "Personal Services"
				If stat_busi_one_type(each_memb) = "06" Then stat_busi_one_type_info(each_memb) = "Paper Route"
				If stat_busi_one_type(each_memb) = "07" Then stat_busi_one_type_info(each_memb) = "In Home Daycare"
				If stat_busi_one_type(each_memb) = "08" Then stat_busi_one_type_info(each_memb) = "Rental Income"
				If stat_busi_one_type(each_memb) = "09" Then stat_busi_one_type_info(each_memb) = "Other"

				stat_busi_one_inc_start_date(each_memb) = replace(stat_busi_one_inc_start_date(each_memb), " ", "/")
				stat_busi_one_inc_end_date(each_memb) = replace(stat_busi_one_inc_end_date(each_memb), " ", "/")
				If stat_busi_one_inc_end_date(each_memb) = "__/__/__" Then stat_busi_one_inc_end_date(each_memb) = ""

				stat_busi_one_method_date(each_memb) = replace(stat_busi_one_method_date(each_memb), " ", "/")
				If stat_busi_one_method_date(each_memb) = "__/__/__" Then stat_busi_one_method_date(each_memb) = ""

				EMReadScreen stat_busi_one_snap_retro_net_inc(each_memb), 8, 10, 55
				EMReadScreen stat_busi_one_snap_prosp_net_inc(each_memb), 8, 10, 69
				stat_busi_one_snap_prosp_net_inc(each_memb) = trim(stat_busi_one_snap_prosp_net_inc(each_memb))

				Call write_value_and_transmit("X", 6, 26)
				EMReadScreen stat_busi_one_snap_retro_gross_inc(each_memb), 8, 11, 43
				EMReadScreen stat_busi_one_snap_retro_expenses(each_memb), 8, 17, 43
				EMReadScreen stat_busi_one_snap_prosp_gross_inc(each_memb), 8, 11, 59
				EMReadScreen stat_busi_one_snap_prosp_expenses(each_memb), 8, 17, 59
				EMReadScreen stat_busi_one_snap_income_verif_code(each_memb), 1, 11, 73
				EMReadScreen stat_busi_one_snap_expense_verif_code(each_memb), 1, 17, 73
				PF3

				If stat_busi_one_snap_income_verif_code(each_memb) = "_" Then stat_busi_one_snap_income_verif_info(each_memb) = ""
				If stat_busi_one_snap_income_verif_code(each_memb) = "1" Then stat_busi_one_snap_income_verif_info(each_memb) = "Income Tax Returns"
				If stat_busi_one_snap_income_verif_code(each_memb) = "2" Then stat_busi_one_snap_income_verif_info(each_memb) = "Receipts of Sales/Purchases"
				If stat_busi_one_snap_income_verif_code(each_memb) = "3" Then stat_busi_one_snap_income_verif_info(each_memb) = "Client Business Records/Ledger"
				If stat_busi_one_snap_income_verif_code(each_memb) = "4" Then stat_busi_one_snap_income_verif_info(each_memb) = "Pending Out of Stat Verifs"
				If stat_busi_one_snap_income_verif_code(each_memb) = "6" Then stat_busi_one_snap_income_verif_info(each_memb) = "Other Document"
				If stat_busi_one_snap_income_verif_code(each_memb) = "N" Then stat_busi_one_snap_income_verif_info(each_memb) = "No Verif Provided"

				If stat_busi_one_snap_expense_verif_code(each_memb) = "_" Then stat_busi_one_snap_expense_verif_info(each_memb) = ""
				If stat_busi_one_snap_expense_verif_code(each_memb) = "1" Then stat_busi_one_snap_expense_verif_info(each_memb) = "Income Tax Returns"
				If stat_busi_one_snap_expense_verif_code(each_memb) = "2" Then stat_busi_one_snap_expense_verif_info(each_memb) = "Receipts of Sales/Purchases"
				If stat_busi_one_snap_expense_verif_code(each_memb) = "3" Then stat_busi_one_snap_expense_verif_info(each_memb) = "Client Business Records/Ledger"
				If stat_busi_one_snap_expense_verif_code(each_memb) = "4" Then stat_busi_one_snap_expense_verif_info(each_memb) = "Pending Out of Stat Verifs"
				If stat_busi_one_snap_expense_verif_code(each_memb) = "6" Then stat_busi_one_snap_expense_verif_info(each_memb) = "Other Document"
				If stat_busi_one_snap_expense_verif_code(each_memb) = "N" Then stat_busi_one_snap_expense_verif_info(each_memb) = "No Verif Provided"
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "02", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_busi_two_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_busi_two_exists(each_memb) = False

			If stat_busi_two_exists(each_memb) = True Then
				stat_busi_two_counted(each_memb) = True
				EMReadScreen stat_busi_two_type(each_memb), 2, 5, 37
				EMReadScreen stat_busi_two_inc_start_date(each_memb), 8, 5, 55
				EMReadScreen stat_busi_two_inc_end_date(each_memb), 8, 5, 72
				EMReadScreen stat_busi_two_method(each_memb), 2, 16, 53
				EMReadScreen stat_busi_two_method_date(each_memb), 8, 16, 63

				If stat_busi_two_type(each_memb) = "01" Then stat_busi_two_type_info(each_memb) = "Farming"
				If stat_busi_two_type(each_memb) = "02" Then stat_busi_two_type_info(each_memb) = "Real Estate"
				If stat_busi_two_type(each_memb) = "03" Then stat_busi_two_type_info(each_memb) = "Home Product Sales"
				If stat_busi_two_type(each_memb) = "04" Then stat_busi_two_type_info(each_memb) = "Other Sales"
				If stat_busi_two_type(each_memb) = "05" Then stat_busi_two_type_info(each_memb) = "Personal Services"
				If stat_busi_two_type(each_memb) = "06" Then stat_busi_two_type_info(each_memb) = "Paper Route"
				If stat_busi_two_type(each_memb) = "07" Then stat_busi_two_type_info(each_memb) = "In Home Daycare"
				If stat_busi_two_type(each_memb) = "08" Then stat_busi_two_type_info(each_memb) = "Rental Income"
				If stat_busi_two_type(each_memb) = "09" Then stat_busi_two_type_info(each_memb) = "Other"

				stat_busi_two_inc_start_date(each_memb) = replace(stat_busi_two_inc_start_date(each_memb), " ", "/")
				stat_busi_two_inc_end_date(each_memb) = replace(stat_busi_two_inc_end_date(each_memb), " ", "/")
				If stat_busi_two_inc_end_date(each_memb) = "__/__/__" Then stat_busi_two_inc_end_date(each_memb) = ""

				stat_busi_two_method_date(each_memb) = replace(stat_busi_two_method_date(each_memb), " ", "/")
				If stat_busi_two_method_date(each_memb) = "__/__/__" Then stat_busi_two_method_date(each_memb) = ""

				EMReadScreen stat_busi_two_snap_retro_net_inc(each_memb), 8, 10, 55
				EMReadScreen stat_busi_two_snap_prosp_net_inc(each_memb), 8, 10, 69
				stat_busi_two_snap_prosp_net_inc(each_memb) = trim(stat_busi_two_snap_prosp_net_inc(each_memb))

				Call write_value_and_transmit("X", 6, 26)
				EMReadScreen stat_busi_two_snap_retro_gross_inc(each_memb), 8, 11, 43
				EMReadScreen stat_busi_two_snap_retro_expenses(each_memb), 8, 17, 43
				EMReadScreen stat_busi_two_snap_prosp_gross_inc(each_memb), 8, 11, 59
				EMReadScreen stat_busi_two_snap_prosp_expenses(each_memb), 8, 17, 59
				EMReadScreen stat_busi_two_snap_income_verif_code(each_memb), 1, 11, 73
				EMReadScreen stat_busi_two_snap_expense_verif_code(each_memb), 1, 17, 73
				PF3

				If stat_busi_two_snap_income_verif_code(each_memb) = "_" Then stat_busi_two_snap_income_verif_info(each_memb) = ""
				If stat_busi_two_snap_income_verif_code(each_memb) = "1" Then stat_busi_two_snap_income_verif_info(each_memb) = "Income Tax Returns"
				If stat_busi_two_snap_income_verif_code(each_memb) = "2" Then stat_busi_two_snap_income_verif_info(each_memb) = "Receipts of Sales/Purchases"
				If stat_busi_two_snap_income_verif_code(each_memb) = "3" Then stat_busi_two_snap_income_verif_info(each_memb) = "Client Business Records/Ledger"
				If stat_busi_two_snap_income_verif_code(each_memb) = "4" Then stat_busi_two_snap_income_verif_info(each_memb) = "Pending Out of Stat Verifs"
				If stat_busi_two_snap_income_verif_code(each_memb) = "6" Then stat_busi_two_snap_income_verif_info(each_memb) = "Other Document"
				If stat_busi_two_snap_income_verif_code(each_memb) = "N" Then stat_busi_two_snap_income_verif_info(each_memb) = "No Verif Provided"

				If stat_busi_two_snap_expense_verif_code(each_memb) = "_" Then stat_busi_two_snap_expense_verif_info(each_memb) = ""
				If stat_busi_two_snap_expense_verif_code(each_memb) = "1" Then stat_busi_two_snap_expense_verif_info(each_memb) = "Income Tax Returns"
				If stat_busi_two_snap_expense_verif_code(each_memb) = "2" Then stat_busi_two_snap_expense_verif_info(each_memb) = "Receipts of Sales/Purchases"
				If stat_busi_two_snap_expense_verif_code(each_memb) = "3" Then stat_busi_two_snap_expense_verif_info(each_memb) = "Client Business Records/Ledger"
				If stat_busi_two_snap_expense_verif_code(each_memb) = "4" Then stat_busi_two_snap_expense_verif_info(each_memb) = "Pending Out of Stat Verifs"
				If stat_busi_two_snap_expense_verif_code(each_memb) = "6" Then stat_busi_two_snap_expense_verif_info(each_memb) = "Other Document"
				If stat_busi_two_snap_expense_verif_code(each_memb) = "N" Then stat_busi_two_snap_expense_verif_info(each_memb) = "No Verif Provided"
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "03", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_busi_three_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_busi_three_exists(each_memb) = False

			If stat_busi_three_exists(each_memb) = True Then
				stat_busi_three_counted(each_memb) = True
				EMReadScreen stat_busi_three_type(each_memb), 2, 5, 37
				EMReadScreen stat_busi_three_inc_start_date(each_memb), 8, 5, 55
				EMReadScreen stat_busi_three_inc_end_date(each_memb), 8, 5, 72
				EMReadScreen stat_busi_three_method(each_memb), 2, 16, 53
				EMReadScreen stat_busi_three_method_date(each_memb), 8, 16, 63

				If stat_busi_three_type(each_memb) = "01" Then stat_busi_three_type_info(each_memb) = "Farming"
				If stat_busi_three_type(each_memb) = "02" Then stat_busi_three_type_info(each_memb) = "Real Estate"
				If stat_busi_three_type(each_memb) = "03" Then stat_busi_three_type_info(each_memb) = "Home Product Sales"
				If stat_busi_three_type(each_memb) = "04" Then stat_busi_three_type_info(each_memb) = "Other Sales"
				If stat_busi_three_type(each_memb) = "05" Then stat_busi_three_type_info(each_memb) = "Personal Services"
				If stat_busi_three_type(each_memb) = "06" Then stat_busi_three_type_info(each_memb) = "Paper Route"
				If stat_busi_three_type(each_memb) = "07" Then stat_busi_three_type_info(each_memb) = "In Home Daycare"
				If stat_busi_three_type(each_memb) = "08" Then stat_busi_three_type_info(each_memb) = "Rental Income"
				If stat_busi_three_type(each_memb) = "09" Then stat_busi_three_type_info(each_memb) = "Other"

				stat_busi_three_inc_start_date(each_memb) = replace(stat_busi_three_inc_start_date(each_memb), " ", "/")
				stat_busi_three_inc_end_date(each_memb) = replace(stat_busi_three_inc_end_date(each_memb), " ", "/")
				If stat_busi_three_inc_end_date(each_memb) = "__/__/__" Then stat_busi_three_inc_end_date(each_memb) = ""

				stat_busi_three_method_date(each_memb) = replace(stat_busi_three_method_date(each_memb), " ", "/")
				If stat_busi_three_method_date(each_memb) = "__/__/__" Then stat_busi_three_method_date(each_memb) = ""

				EMReadScreen stat_busi_three_snap_retro_net_inc(each_memb), 8, 10, 55
				EMReadScreen stat_busi_three_snap_prosp_net_inc(each_memb), 8, 10, 69
				stat_busi_three_snap_prosp_net_inc(each_memb) = trim(stat_busi_three_snap_prosp_net_inc(each_memb))

				Call write_value_and_transmit("X", 6, 26)
				EMReadScreen stat_busi_three_snap_retro_gross_inc(each_memb), 8, 11, 43
				EMReadScreen stat_busi_three_snap_retro_expenses(each_memb), 8, 17, 43
				EMReadScreen stat_busi_three_snap_prosp_gross_inc(each_memb), 8, 11, 59
				EMReadScreen stat_busi_three_snap_prosp_expenses(each_memb), 8, 17, 59
				EMReadScreen stat_busi_three_snap_income_verif_code(each_memb), 1, 11, 73
				EMReadScreen stat_busi_three_snap_expense_verif_code(each_memb), 1, 17, 73
				PF3

				If stat_busi_three_snap_income_verif_code(each_memb) = "_" Then stat_busi_three_snap_income_verif_info(each_memb) = ""
				If stat_busi_three_snap_income_verif_code(each_memb) = "1" Then stat_busi_three_snap_income_verif_info(each_memb) = "Income Tax Returns"
				If stat_busi_three_snap_income_verif_code(each_memb) = "2" Then stat_busi_three_snap_income_verif_info(each_memb) = "Receipts of Sales/Purchases"
				If stat_busi_three_snap_income_verif_code(each_memb) = "3" Then stat_busi_three_snap_income_verif_info(each_memb) = "Client Business Records/Ledger"
				If stat_busi_three_snap_income_verif_code(each_memb) = "4" Then stat_busi_three_snap_income_verif_info(each_memb) = "Pending Out of Stat Verifs"
				If stat_busi_three_snap_income_verif_code(each_memb) = "6" Then stat_busi_three_snap_income_verif_info(each_memb) = "Other Document"
				If stat_busi_three_snap_income_verif_code(each_memb) = "N" Then stat_busi_three_snap_income_verif_info(each_memb) = "No Verif Provided"

				If stat_busi_three_snap_expense_verif_code(each_memb) = "_" Then stat_busi_three_snap_expense_verif_info(each_memb) = ""
				If stat_busi_three_snap_expense_verif_code(each_memb) = "1" Then stat_busi_three_snap_expense_verif_info(each_memb) = "Income Tax Returns"
				If stat_busi_three_snap_expense_verif_code(each_memb) = "2" Then stat_busi_three_snap_expense_verif_info(each_memb) = "Receipts of Sales/Purchases"
				If stat_busi_three_snap_expense_verif_code(each_memb) = "3" Then stat_busi_three_snap_expense_verif_info(each_memb) = "Client Business Records/Ledger"
				If stat_busi_three_snap_expense_verif_code(each_memb) = "4" Then stat_busi_three_snap_expense_verif_info(each_memb) = "Pending Out of Stat Verifs"
				If stat_busi_three_snap_expense_verif_code(each_memb) = "6" Then stat_busi_three_snap_expense_verif_info(each_memb) = "Other Document"
				If stat_busi_three_snap_expense_verif_code(each_memb) = "N" Then stat_busi_three_snap_expense_verif_info(each_memb) = "No Verif Provided"
			End If

		Next

		call navigate_to_MAXIS_screen("STAT", "UNEA")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_unea_one_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_unea_one_exists(each_memb) = False

			If stat_unea_one_exists(each_memb) = True Then
				EMReadScreen stat_unea_one_type_code(each_memb), 2, 5, 37
				stat_unea_one_counted(each_memb) = True
				If stat_unea_one_type_code(each_memb) = "01" Then stat_unea_one_type_info(each_memb) = "RSDI, Disability"
				If stat_unea_one_type_code(each_memb) = "02" Then stat_unea_one_type_info(each_memb) = "RSDI, No Disability"
				If stat_unea_one_type_code(each_memb) = "06" Then stat_unea_one_type_info(each_memb) = "SSI"
				If stat_unea_one_type_code(each_memb) = "03" Then stat_unea_one_type_info(each_memb) = "Non-MN Public Assistance"
				If stat_unea_one_type_code(each_memb) = "11" Then stat_unea_one_type_info(each_memb) = "VA Disability Benefit"
				If stat_unea_one_type_code(each_memb) = "12" Then stat_unea_one_type_info(each_memb) = "VA Pension"
				If stat_unea_one_type_code(each_memb) = "13" Then stat_unea_one_type_info(each_memb) = "VA other"
				If stat_unea_one_type_code(each_memb) = "38" Then stat_unea_one_type_info(each_memb) = "VA Aid & Attendance"
				If stat_unea_one_type_code(each_memb) = "14" Then stat_unea_one_type_info(each_memb) = "Unemployment Insurance"
				If stat_unea_one_type_code(each_memb) = "15" Then stat_unea_one_type_info(each_memb) = "Worker's Comp"
				If stat_unea_one_type_code(each_memb) = "16" Then stat_unea_one_type_info(each_memb) = "Railroad Retirement"
				If stat_unea_one_type_code(each_memb) = "17" Then stat_unea_one_type_info(each_memb) = "Other Retirement"
				If stat_unea_one_type_code(each_memb) = "18" Then stat_unea_one_type_info(each_memb) = "Military Entitlement"
				If stat_unea_one_type_code(each_memb) = "19" Then stat_unea_one_type_info(each_memb) = "Foster Care Child Requesting SNAP"
				If stat_unea_one_type_code(each_memb) = "20" Then stat_unea_one_type_info(each_memb) = "Foster Care Child NOT Requesting SNAP"
				If stat_unea_one_type_code(each_memb) = "21" Then stat_unea_one_type_info(each_memb) = "Foster Care Adult Requesting SNAP"
				If stat_unea_one_type_code(each_memb) = "22" Then stat_unea_one_type_info(each_memb) = "Foster Care Adult NOT Requesting SNAP"
				If stat_unea_one_type_code(each_memb) = "23" Then stat_unea_one_type_info(each_memb) = "Dividends"
				If stat_unea_one_type_code(each_memb) = "24" Then stat_unea_one_type_info(each_memb) = "Interest"
				If stat_unea_one_type_code(each_memb) = "25" Then stat_unea_one_type_info(each_memb) = "Counted Gifts or Prizes"
				If stat_unea_one_type_code(each_memb) = "26" Then stat_unea_one_type_info(each_memb) = "Strike Benefit"
				If stat_unea_one_type_code(each_memb) = "27" Then stat_unea_one_type_info(each_memb) = "Contract for Deed"
				If stat_unea_one_type_code(each_memb) = "28" Then stat_unea_one_type_info(each_memb) = "Illegal Income"
				If stat_unea_one_type_code(each_memb) = "29" Then stat_unea_one_type_info(each_memb) = "Other Countable"
				If stat_unea_one_type_code(each_memb) = "30" Then stat_unea_one_type_info(each_memb) = "Infrequent, <30, Not Counted"
				If stat_unea_one_type_code(each_memb) = "31" Then stat_unea_one_type_info(each_memb) = "Other SNAP Only"
				If stat_unea_one_type_code(each_memb) = "08" Then stat_unea_one_type_info(each_memb) = "Direct Child Support"
				If stat_unea_one_type_code(each_memb) = "35" Then stat_unea_one_type_info(each_memb) = "Direct Spousal Support"
				If stat_unea_one_type_code(each_memb) = "36" Then stat_unea_one_type_info(each_memb) = "Disbursed Child Support"
				If stat_unea_one_type_code(each_memb) = "37" Then stat_unea_one_type_info(each_memb) = "Disbursed Spousal Support"
				If stat_unea_one_type_code(each_memb) = "39" Then stat_unea_one_type_info(each_memb) = "Disbursed Child Support Arrears"
				If stat_unea_one_type_code(each_memb) = "40" Then stat_unea_one_type_info(each_memb) = "Disbursed Spousal Support Arrears"
				If stat_unea_one_type_code(each_memb) = "43" Then stat_unea_one_type_info(each_memb) = "Disbursed Excess Child Support"
				If stat_unea_one_type_code(each_memb) = "44" Then stat_unea_one_type_info(each_memb) = "MSA - Excess Income for SSI"
				If stat_unea_one_type_code(each_memb) = "45" Then stat_unea_one_type_info(each_memb) = "County 88 Child Support"
				If stat_unea_one_type_code(each_memb) = "46" Then stat_unea_one_type_info(each_memb) = "County 88 Gaming"
				If stat_unea_one_type_code(each_memb) = "47" Then stat_unea_one_type_info(each_memb) = "Counted Tribal Income"
				If stat_unea_one_type_code(each_memb) = "48" Then stat_unea_one_type_info(each_memb) = "Trust income"
				If stat_unea_one_type_code(each_memb) = "49" Then stat_unea_one_type_info(each_memb) = "Non-Recurring Income > $60 per Quarter"
				EMReadScreen stat_unea_one_verif_code(each_memb), 1, 5, 65
				If stat_unea_one_verif_code(each_memb) = "1" Then stat_unea_one_verif_info(each_memb) = "Copy of Checks"
				If stat_unea_one_verif_code(each_memb) = "2" Then stat_unea_one_verif_info(each_memb) = "Award Letter"
				If stat_unea_one_verif_code(each_memb) = "3" Then stat_unea_one_verif_info(each_memb) = "System Initiated Verif"
				If stat_unea_one_verif_code(each_memb) = "4" Then stat_unea_one_verif_info(each_memb) = "Collateral Statement"
				If stat_unea_one_verif_code(each_memb) = "5" Then stat_unea_one_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_unea_one_verif_code(each_memb) = "6" Then stat_unea_one_verif_info(each_memb) = "Other Document"
				If stat_unea_one_verif_code(each_memb) = "7" Then stat_unea_one_verif_info(each_memb) = "Worker Initiated Verif"
				If stat_unea_one_verif_code(each_memb) = "8" Then stat_unea_one_verif_info(each_memb) = "RI Stubs"
				If stat_unea_one_verif_code(each_memb) = "N" Then stat_unea_one_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_unea_one_inc_start_date(each_memb), 8, 7, 37
				EMReadScreen stat_unea_one_inc_end_date(each_memb), 8, 7, 68

				stat_unea_one_inc_start_date(each_memb) = replace(stat_unea_one_inc_start_date(each_memb), " ", "/")
				stat_unea_one_inc_end_date(each_memb) = replace(stat_unea_one_inc_end_date(each_memb), " ", "/")
				iF stat_unea_one_inc_end_date(each_memb) = "__/__/__" Then stat_unea_one_inc_end_date(each_memb) = ""

				Call write_value_and_transmit("X", 10, 26)
				EMReadScreen stat_unea_one_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_unea_one_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 52
				EMReadScreen stat_unea_one_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 52

				If stat_unea_one_snap_pic_pay_freq(each_memb) = "_" Then stat_unea_one_snap_pic_pay_freq(each_memb) = ""
				If stat_unea_one_snap_pic_pay_freq(each_memb) = "1" Then stat_unea_one_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_unea_one_snap_pic_pay_freq(each_memb) = "2" Then stat_unea_one_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_unea_one_snap_pic_pay_freq(each_memb) = "3" Then stat_unea_one_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_unea_one_snap_pic_pay_freq(each_memb) = "4" Then stat_unea_one_snap_pic_pay_freq(each_memb) = "Weekly"
				stat_unea_one_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_unea_one_snap_pic_ave_inc_per_pay(each_memb))
				stat_unea_one_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_unea_one_snap_pic_prosp_monthly_inc(each_memb))
				PF3

			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "02", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_unea_two_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_unea_two_exists(each_memb) = False

			If stat_unea_two_exists(each_memb) = True Then
				EMReadScreen stat_unea_two_type_code(each_memb), 2, 5, 37
				stat_unea_two_counted(each_memb) = True
				If stat_unea_two_type_code(each_memb) = "01" Then stat_unea_two_type_info(each_memb) = "RSDI, Disability"
				If stat_unea_two_type_code(each_memb) = "02" Then stat_unea_two_type_info(each_memb) = "RSDI, No Disability"
				If stat_unea_two_type_code(each_memb) = "06" Then stat_unea_two_type_info(each_memb) = "SSI"
				If stat_unea_two_type_code(each_memb) = "03" Then stat_unea_two_type_info(each_memb) = "Non-MN Public Assistance"
				If stat_unea_two_type_code(each_memb) = "11" Then stat_unea_two_type_info(each_memb) = "VA Disability Benefit"
				If stat_unea_two_type_code(each_memb) = "12" Then stat_unea_two_type_info(each_memb) = "VA Pension"
				If stat_unea_two_type_code(each_memb) = "13" Then stat_unea_two_type_info(each_memb) = "VA other"
				If stat_unea_two_type_code(each_memb) = "38" Then stat_unea_two_type_info(each_memb) = "VA Aid & Attendance"
				If stat_unea_two_type_code(each_memb) = "14" Then stat_unea_two_type_info(each_memb) = "Unemployment Insurance"
				If stat_unea_two_type_code(each_memb) = "15" Then stat_unea_two_type_info(each_memb) = "Worker's Comp"
				If stat_unea_two_type_code(each_memb) = "16" Then stat_unea_two_type_info(each_memb) = "Railroad Retirement"
				If stat_unea_two_type_code(each_memb) = "17" Then stat_unea_two_type_info(each_memb) = "Other Retirement"
				If stat_unea_two_type_code(each_memb) = "18" Then stat_unea_two_type_info(each_memb) = "Military Entitlement"
				If stat_unea_two_type_code(each_memb) = "19" Then stat_unea_two_type_info(each_memb) = "Foster Care Child Requesting SNAP"
				If stat_unea_two_type_code(each_memb) = "20" Then stat_unea_two_type_info(each_memb) = "Foster Care Child NOT Requesting SNAP"
				If stat_unea_two_type_code(each_memb) = "21" Then stat_unea_two_type_info(each_memb) = "Foster Care Adult Requesting SNAP"
				If stat_unea_two_type_code(each_memb) = "22" Then stat_unea_two_type_info(each_memb) = "Foster Care Adult NOT Requesting SNAP"
				If stat_unea_two_type_code(each_memb) = "23" Then stat_unea_two_type_info(each_memb) = "Dividends"
				If stat_unea_two_type_code(each_memb) = "24" Then stat_unea_two_type_info(each_memb) = "Interest"
				If stat_unea_two_type_code(each_memb) = "25" Then stat_unea_two_type_info(each_memb) = "Counted Gifts or Prizes"
				If stat_unea_two_type_code(each_memb) = "26" Then stat_unea_two_type_info(each_memb) = "Strike Benefit"
				If stat_unea_two_type_code(each_memb) = "27" Then stat_unea_two_type_info(each_memb) = "Contract for Deed"
				If stat_unea_two_type_code(each_memb) = "28" Then stat_unea_two_type_info(each_memb) = "Illegal Income"
				If stat_unea_two_type_code(each_memb) = "29" Then stat_unea_two_type_info(each_memb) = "Other Countable"
				If stat_unea_two_type_code(each_memb) = "30" Then stat_unea_two_type_info(each_memb) = "Infrequent, <30, Not Counted"
				If stat_unea_two_type_code(each_memb) = "31" Then stat_unea_two_type_info(each_memb) = "Other SNAP Only"
				If stat_unea_two_type_code(each_memb) = "08" Then stat_unea_two_type_info(each_memb) = "Direct Child Support"
				If stat_unea_two_type_code(each_memb) = "35" Then stat_unea_two_type_info(each_memb) = "Direct Spousal Support"
				If stat_unea_two_type_code(each_memb) = "36" Then stat_unea_two_type_info(each_memb) = "Disbursed Child Support"
				If stat_unea_two_type_code(each_memb) = "37" Then stat_unea_two_type_info(each_memb) = "Disbursed Spousal Support"
				If stat_unea_two_type_code(each_memb) = "39" Then stat_unea_two_type_info(each_memb) = "Disbursed Child Support Arrears"
				If stat_unea_two_type_code(each_memb) = "40" Then stat_unea_two_type_info(each_memb) = "Disbursed Spousal Support Arrears"
				If stat_unea_two_type_code(each_memb) = "43" Then stat_unea_two_type_info(each_memb) = "Disbursed Excess Child Support"
				If stat_unea_two_type_code(each_memb) = "44" Then stat_unea_two_type_info(each_memb) = "MSA - Excess Income for SSI"
				If stat_unea_two_type_code(each_memb) = "45" Then stat_unea_two_type_info(each_memb) = "County 88 Child Support"
				If stat_unea_two_type_code(each_memb) = "46" Then stat_unea_two_type_info(each_memb) = "County 88 Gaming"
				If stat_unea_two_type_code(each_memb) = "47" Then stat_unea_two_type_info(each_memb) = "Counted Tribal Income"
				If stat_unea_two_type_code(each_memb) = "48" Then stat_unea_two_type_info(each_memb) = "Trust income"
				If stat_unea_two_type_code(each_memb) = "49" Then stat_unea_two_type_info(each_memb) = "Non-Recurring Income > $60 per Quarter"
				EMReadScreen stat_unea_two_verif_code(each_memb), 1, 5, 65
				If stat_unea_two_verif_code(each_memb) = "1" Then stat_unea_two_verif_info(each_memb) = "Copy of Checks"
				If stat_unea_two_verif_code(each_memb) = "2" Then stat_unea_two_verif_info(each_memb) = "Award Letter"
				If stat_unea_two_verif_code(each_memb) = "3" Then stat_unea_two_verif_info(each_memb) = "System Initiated Verif"
				If stat_unea_two_verif_code(each_memb) = "4" Then stat_unea_two_verif_info(each_memb) = "Collateral Statement"
				If stat_unea_two_verif_code(each_memb) = "5" Then stat_unea_two_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_unea_two_verif_code(each_memb) = "6" Then stat_unea_two_verif_info(each_memb) = "Other Document"
				If stat_unea_two_verif_code(each_memb) = "7" Then stat_unea_two_verif_info(each_memb) = "Worker Initiated Verif"
				If stat_unea_two_verif_code(each_memb) = "8" Then stat_unea_two_verif_info(each_memb) = "RI Stubs"
				If stat_unea_two_verif_code(each_memb) = "N" Then stat_unea_two_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_unea_two_inc_start_date(each_memb), 8, 7, 37
				EMReadScreen stat_unea_two_inc_end_date(each_memb), 8, 7, 68

				stat_unea_two_inc_start_date(each_memb) = replace(stat_unea_two_inc_start_date(each_memb), " ", "/")
				stat_unea_two_inc_end_date(each_memb) = replace(stat_unea_two_inc_end_date(each_memb), " ", "/")
				iF stat_unea_two_inc_end_date(each_memb) = "__/__/__" Then stat_unea_two_inc_end_date(each_memb) = ""

				Call write_value_and_transmit("X", 10, 26)
				EMReadScreen stat_unea_two_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_unea_two_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 52
				EMReadScreen stat_unea_two_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 52

				If stat_unea_two_snap_pic_pay_freq(each_memb) = "_" Then stat_unea_two_snap_pic_pay_freq(each_memb) = ""
				If stat_unea_two_snap_pic_pay_freq(each_memb) = "1" Then stat_unea_two_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_unea_two_snap_pic_pay_freq(each_memb) = "2" Then stat_unea_two_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_unea_two_snap_pic_pay_freq(each_memb) = "3" Then stat_unea_two_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_unea_two_snap_pic_pay_freq(each_memb) = "4" Then stat_unea_two_snap_pic_pay_freq(each_memb) = "Weekly"
				stat_unea_two_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_unea_two_snap_pic_ave_inc_per_pay(each_memb))
				stat_unea_two_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_unea_two_snap_pic_prosp_monthly_inc(each_memb))
				PF3

			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "03", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_unea_three_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_unea_three_exists(each_memb) = False

			If stat_unea_three_exists(each_memb) = True Then
				EMReadScreen stat_unea_three_type_code(each_memb), 2, 5, 37
				stat_unea_three_counted(each_memb) = True
				If stat_unea_three_type_code(each_memb) = "01" Then stat_unea_three_type_info(each_memb) = "RSDI, Disability"
				If stat_unea_three_type_code(each_memb) = "02" Then stat_unea_three_type_info(each_memb) = "RSDI, No Disability"
				If stat_unea_three_type_code(each_memb) = "06" Then stat_unea_three_type_info(each_memb) = "SSI"
				If stat_unea_three_type_code(each_memb) = "03" Then stat_unea_three_type_info(each_memb) = "Non-MN Public Assistance"
				If stat_unea_three_type_code(each_memb) = "11" Then stat_unea_three_type_info(each_memb) = "VA Disability Benefit"
				If stat_unea_three_type_code(each_memb) = "12" Then stat_unea_three_type_info(each_memb) = "VA Pension"
				If stat_unea_three_type_code(each_memb) = "13" Then stat_unea_three_type_info(each_memb) = "VA other"
				If stat_unea_three_type_code(each_memb) = "38" Then stat_unea_three_type_info(each_memb) = "VA Aid & Attendance"
				If stat_unea_three_type_code(each_memb) = "14" Then stat_unea_three_type_info(each_memb) = "Unemployment Insurance"
				If stat_unea_three_type_code(each_memb) = "15" Then stat_unea_three_type_info(each_memb) = "Worker's Comp"
				If stat_unea_three_type_code(each_memb) = "16" Then stat_unea_three_type_info(each_memb) = "Railroad Retirement"
				If stat_unea_three_type_code(each_memb) = "17" Then stat_unea_three_type_info(each_memb) = "Other Retirement"
				If stat_unea_three_type_code(each_memb) = "18" Then stat_unea_three_type_info(each_memb) = "Military Entitlement"
				If stat_unea_three_type_code(each_memb) = "19" Then stat_unea_three_type_info(each_memb) = "Foster Care Child Requesting SNAP"
				If stat_unea_three_type_code(each_memb) = "20" Then stat_unea_three_type_info(each_memb) = "Foster Care Child NOT Requesting SNAP"
				If stat_unea_three_type_code(each_memb) = "21" Then stat_unea_three_type_info(each_memb) = "Foster Care Adult Requesting SNAP"
				If stat_unea_three_type_code(each_memb) = "22" Then stat_unea_three_type_info(each_memb) = "Foster Care Adult NOT Requesting SNAP"
				If stat_unea_three_type_code(each_memb) = "23" Then stat_unea_three_type_info(each_memb) = "Dividends"
				If stat_unea_three_type_code(each_memb) = "24" Then stat_unea_three_type_info(each_memb) = "Interest"
				If stat_unea_three_type_code(each_memb) = "25" Then stat_unea_three_type_info(each_memb) = "Counted Gifts or Prizes"
				If stat_unea_three_type_code(each_memb) = "26" Then stat_unea_three_type_info(each_memb) = "Strike Benefit"
				If stat_unea_three_type_code(each_memb) = "27" Then stat_unea_three_type_info(each_memb) = "Contract for Deed"
				If stat_unea_three_type_code(each_memb) = "28" Then stat_unea_three_type_info(each_memb) = "Illegal Income"
				If stat_unea_three_type_code(each_memb) = "29" Then stat_unea_three_type_info(each_memb) = "Other Countable"
				If stat_unea_three_type_code(each_memb) = "30" Then stat_unea_three_type_info(each_memb) = "Infrequent, <30, Not Counted"
				If stat_unea_three_type_code(each_memb) = "31" Then stat_unea_three_type_info(each_memb) = "Other SNAP Only"
				If stat_unea_three_type_code(each_memb) = "08" Then stat_unea_three_type_info(each_memb) = "Direct Child Support"
				If stat_unea_three_type_code(each_memb) = "35" Then stat_unea_three_type_info(each_memb) = "Direct Spousal Support"
				If stat_unea_three_type_code(each_memb) = "36" Then stat_unea_three_type_info(each_memb) = "Disbursed Child Support"
				If stat_unea_three_type_code(each_memb) = "37" Then stat_unea_three_type_info(each_memb) = "Disbursed Spousal Support"
				If stat_unea_three_type_code(each_memb) = "39" Then stat_unea_three_type_info(each_memb) = "Disbursed Child Support Arrears"
				If stat_unea_three_type_code(each_memb) = "40" Then stat_unea_three_type_info(each_memb) = "Disbursed Spousal Support Arrears"
				If stat_unea_three_type_code(each_memb) = "43" Then stat_unea_three_type_info(each_memb) = "Disbursed Excess Child Support"
				If stat_unea_three_type_code(each_memb) = "44" Then stat_unea_three_type_info(each_memb) = "MSA - Excess Income for SSI"
				If stat_unea_three_type_code(each_memb) = "45" Then stat_unea_three_type_info(each_memb) = "County 88 Child Support"
				If stat_unea_three_type_code(each_memb) = "46" Then stat_unea_three_type_info(each_memb) = "County 88 Gaming"
				If stat_unea_three_type_code(each_memb) = "47" Then stat_unea_three_type_info(each_memb) = "Counted Tribal Income"
				If stat_unea_three_type_code(each_memb) = "48" Then stat_unea_three_type_info(each_memb) = "Trust income"
				If stat_unea_three_type_code(each_memb) = "49" Then stat_unea_three_type_info(each_memb) = "Non-Recurring Income > $60 per Quarter"
				EMReadScreen stat_unea_three_verif_code(each_memb), 1, 5, 65
				If stat_unea_three_verif_code(each_memb) = "1" Then stat_unea_three_verif_info(each_memb) = "Copy of Checks"
				If stat_unea_three_verif_code(each_memb) = "2" Then stat_unea_three_verif_info(each_memb) = "Award Letter"
				If stat_unea_three_verif_code(each_memb) = "3" Then stat_unea_three_verif_info(each_memb) = "System Initiated Verif"
				If stat_unea_three_verif_code(each_memb) = "4" Then stat_unea_three_verif_info(each_memb) = "Collateral Statement"
				If stat_unea_three_verif_code(each_memb) = "5" Then stat_unea_three_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_unea_three_verif_code(each_memb) = "6" Then stat_unea_three_verif_info(each_memb) = "Other Document"
				If stat_unea_three_verif_code(each_memb) = "7" Then stat_unea_three_verif_info(each_memb) = "Worker Initiated Verif"
				If stat_unea_three_verif_code(each_memb) = "8" Then stat_unea_three_verif_info(each_memb) = "RI Stubs"
				If stat_unea_three_verif_code(each_memb) = "N" Then stat_unea_three_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_unea_three_inc_start_date(each_memb), 8, 7, 37
				EMReadScreen stat_unea_three_inc_end_date(each_memb), 8, 7, 68

				stat_unea_three_inc_start_date(each_memb) = replace(stat_unea_three_inc_start_date(each_memb), " ", "/")
				stat_unea_three_inc_end_date(each_memb) = replace(stat_unea_three_inc_end_date(each_memb), " ", "/")
				iF stat_unea_three_inc_end_date(each_memb) = "__/__/__" Then stat_unea_three_inc_end_date(each_memb) = ""

				Call write_value_and_transmit("X", 10, 26)
				EMReadScreen stat_unea_three_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_unea_three_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 52
				EMReadScreen stat_unea_three_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 52

				If stat_unea_three_snap_pic_pay_freq(each_memb) = "_" Then stat_unea_three_snap_pic_pay_freq(each_memb) = ""
				If stat_unea_three_snap_pic_pay_freq(each_memb) = "1" Then stat_unea_three_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_unea_three_snap_pic_pay_freq(each_memb) = "2" Then stat_unea_three_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_unea_three_snap_pic_pay_freq(each_memb) = "3" Then stat_unea_three_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_unea_three_snap_pic_pay_freq(each_memb) = "4" Then stat_unea_three_snap_pic_pay_freq(each_memb) = "Weekly"
				stat_unea_three_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_unea_three_snap_pic_ave_inc_per_pay(each_memb))
				stat_unea_three_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_unea_three_snap_pic_prosp_monthly_inc(each_memb))
				PF3

			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "04", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_unea_four_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_unea_four_exists(each_memb) = False

			If stat_unea_four_exists(each_memb) = True Then
				EMReadScreen stat_unea_four_type_code(each_memb), 2, 5, 37
				stat_unea_four_counted(each_memb) = True
				If stat_unea_four_type_code(each_memb) = "01" Then stat_unea_four_type_info(each_memb) = "RSDI, Disability"
				If stat_unea_four_type_code(each_memb) = "02" Then stat_unea_four_type_info(each_memb) = "RSDI, No Disability"
				If stat_unea_four_type_code(each_memb) = "06" Then stat_unea_four_type_info(each_memb) = "SSI"
				If stat_unea_four_type_code(each_memb) = "03" Then stat_unea_four_type_info(each_memb) = "Non-MN Public Assistance"
				If stat_unea_four_type_code(each_memb) = "11" Then stat_unea_four_type_info(each_memb) = "VA Disability Benefit"
				If stat_unea_four_type_code(each_memb) = "12" Then stat_unea_four_type_info(each_memb) = "VA Pension"
				If stat_unea_four_type_code(each_memb) = "13" Then stat_unea_four_type_info(each_memb) = "VA other"
				If stat_unea_four_type_code(each_memb) = "38" Then stat_unea_four_type_info(each_memb) = "VA Aid & Attendance"
				If stat_unea_four_type_code(each_memb) = "14" Then stat_unea_four_type_info(each_memb) = "Unemployment Insurance"
				If stat_unea_four_type_code(each_memb) = "15" Then stat_unea_four_type_info(each_memb) = "Worker's Comp"
				If stat_unea_four_type_code(each_memb) = "16" Then stat_unea_four_type_info(each_memb) = "Railroad Retirement"
				If stat_unea_four_type_code(each_memb) = "17" Then stat_unea_four_type_info(each_memb) = "Other Retirement"
				If stat_unea_four_type_code(each_memb) = "18" Then stat_unea_four_type_info(each_memb) = "Military Entitlement"
				If stat_unea_four_type_code(each_memb) = "19" Then stat_unea_four_type_info(each_memb) = "Foster Care Child Requesting SNAP"
				If stat_unea_four_type_code(each_memb) = "20" Then stat_unea_four_type_info(each_memb) = "Foster Care Child NOT Requesting SNAP"
				If stat_unea_four_type_code(each_memb) = "21" Then stat_unea_four_type_info(each_memb) = "Foster Care Adult Requesting SNAP"
				If stat_unea_four_type_code(each_memb) = "22" Then stat_unea_four_type_info(each_memb) = "Foster Care Adult NOT Requesting SNAP"
				If stat_unea_four_type_code(each_memb) = "23" Then stat_unea_four_type_info(each_memb) = "Dividends"
				If stat_unea_four_type_code(each_memb) = "24" Then stat_unea_four_type_info(each_memb) = "Interest"
				If stat_unea_four_type_code(each_memb) = "25" Then stat_unea_four_type_info(each_memb) = "Counted Gifts or Prizes"
				If stat_unea_four_type_code(each_memb) = "26" Then stat_unea_four_type_info(each_memb) = "Strike Benefit"
				If stat_unea_four_type_code(each_memb) = "27" Then stat_unea_four_type_info(each_memb) = "Contract for Deed"
				If stat_unea_four_type_code(each_memb) = "28" Then stat_unea_four_type_info(each_memb) = "Illegal Income"
				If stat_unea_four_type_code(each_memb) = "29" Then stat_unea_four_type_info(each_memb) = "Other Countable"
				If stat_unea_four_type_code(each_memb) = "30" Then stat_unea_four_type_info(each_memb) = "Infrequent, <30, Not Counted"
				If stat_unea_four_type_code(each_memb) = "31" Then stat_unea_four_type_info(each_memb) = "Other SNAP Only"
				If stat_unea_four_type_code(each_memb) = "08" Then stat_unea_four_type_info(each_memb) = "Direct Child Support"
				If stat_unea_four_type_code(each_memb) = "35" Then stat_unea_four_type_info(each_memb) = "Direct Spousal Support"
				If stat_unea_four_type_code(each_memb) = "36" Then stat_unea_four_type_info(each_memb) = "Disbursed Child Support"
				If stat_unea_four_type_code(each_memb) = "37" Then stat_unea_four_type_info(each_memb) = "Disbursed Spousal Support"
				If stat_unea_four_type_code(each_memb) = "39" Then stat_unea_four_type_info(each_memb) = "Disbursed Child Support Arrears"
				If stat_unea_four_type_code(each_memb) = "40" Then stat_unea_four_type_info(each_memb) = "Disbursed Spousal Support Arrears"
				If stat_unea_four_type_code(each_memb) = "43" Then stat_unea_four_type_info(each_memb) = "Disbursed Excess Child Support"
				If stat_unea_four_type_code(each_memb) = "44" Then stat_unea_four_type_info(each_memb) = "MSA - Excess Income for SSI"
				If stat_unea_four_type_code(each_memb) = "45" Then stat_unea_four_type_info(each_memb) = "County 88 Child Support"
				If stat_unea_four_type_code(each_memb) = "46" Then stat_unea_four_type_info(each_memb) = "County 88 Gaming"
				If stat_unea_four_type_code(each_memb) = "47" Then stat_unea_four_type_info(each_memb) = "Counted Tribal Income"
				If stat_unea_four_type_code(each_memb) = "48" Then stat_unea_four_type_info(each_memb) = "Trust income"
				If stat_unea_four_type_code(each_memb) = "49" Then stat_unea_four_type_info(each_memb) = "Non-Recurring Income > $60 per Quarter"
				EMReadScreen stat_unea_four_verif_code(each_memb), 1, 5, 65
				If stat_unea_four_verif_code(each_memb) = "1" Then stat_unea_four_verif_info(each_memb) = "Copy of Checks"
				If stat_unea_four_verif_code(each_memb) = "2" Then stat_unea_four_verif_info(each_memb) = "Award Letter"
				If stat_unea_four_verif_code(each_memb) = "3" Then stat_unea_four_verif_info(each_memb) = "System Initiated Verif"
				If stat_unea_four_verif_code(each_memb) = "4" Then stat_unea_four_verif_info(each_memb) = "Collateral Statement"
				If stat_unea_four_verif_code(each_memb) = "5" Then stat_unea_four_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_unea_four_verif_code(each_memb) = "6" Then stat_unea_four_verif_info(each_memb) = "Other Document"
				If stat_unea_four_verif_code(each_memb) = "7" Then stat_unea_four_verif_info(each_memb) = "Worker Initiated Verif"
				If stat_unea_four_verif_code(each_memb) = "8" Then stat_unea_four_verif_info(each_memb) = "RI Stubs"
				If stat_unea_four_verif_code(each_memb) = "N" Then stat_unea_four_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_unea_four_inc_start_date(each_memb), 8, 7, 37
				EMReadScreen stat_unea_four_inc_end_date(each_memb), 8, 7, 68

				stat_unea_four_inc_start_date(each_memb) = replace(stat_unea_four_inc_start_date(each_memb), " ", "/")
				stat_unea_four_inc_end_date(each_memb) = replace(stat_unea_four_inc_end_date(each_memb), " ", "/")
				iF stat_unea_four_inc_end_date(each_memb) = "__/__/__" Then stat_unea_four_inc_end_date(each_memb) = ""

				Call write_value_and_transmit("X", 10, 26)
				EMReadScreen stat_unea_four_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_unea_four_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 52
				EMReadScreen stat_unea_four_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 52

				If stat_unea_four_snap_pic_pay_freq(each_memb) = "_" Then stat_unea_four_snap_pic_pay_freq(each_memb) = ""
				If stat_unea_four_snap_pic_pay_freq(each_memb) = "1" Then stat_unea_four_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_unea_four_snap_pic_pay_freq(each_memb) = "2" Then stat_unea_four_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_unea_four_snap_pic_pay_freq(each_memb) = "3" Then stat_unea_four_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_unea_four_snap_pic_pay_freq(each_memb) = "4" Then stat_unea_four_snap_pic_pay_freq(each_memb) = "Weekly"
				stat_unea_four_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_unea_four_snap_pic_ave_inc_per_pay(each_memb))
				stat_unea_four_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_unea_four_snap_pic_prosp_monthly_inc(each_memb))
				PF3

			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "05", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_unea_five_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_unea_five_exists(each_memb) = False

			If stat_unea_five_exists(each_memb) = True Then
				EMReadScreen stat_unea_five_type_code(each_memb), 2, 5, 37
				stat_unea_five_counted(each_memb) = True
				If stat_unea_five_type_code(each_memb) = "01" Then stat_unea_five_type_info(each_memb) = "RSDI, Disability"
				If stat_unea_five_type_code(each_memb) = "02" Then stat_unea_five_type_info(each_memb) = "RSDI, No Disability"
				If stat_unea_five_type_code(each_memb) = "06" Then stat_unea_five_type_info(each_memb) = "SSI"
				If stat_unea_five_type_code(each_memb) = "03" Then stat_unea_five_type_info(each_memb) = "Non-MN Public Assistance"
				If stat_unea_five_type_code(each_memb) = "11" Then stat_unea_five_type_info(each_memb) = "VA Disability Benefit"
				If stat_unea_five_type_code(each_memb) = "12" Then stat_unea_five_type_info(each_memb) = "VA Pension"
				If stat_unea_five_type_code(each_memb) = "13" Then stat_unea_five_type_info(each_memb) = "VA other"
				If stat_unea_five_type_code(each_memb) = "38" Then stat_unea_five_type_info(each_memb) = "VA Aid & Attendance"
				If stat_unea_five_type_code(each_memb) = "14" Then stat_unea_five_type_info(each_memb) = "Unemployment Insurance"
				If stat_unea_five_type_code(each_memb) = "15" Then stat_unea_five_type_info(each_memb) = "Worker's Comp"
				If stat_unea_five_type_code(each_memb) = "16" Then stat_unea_five_type_info(each_memb) = "Railroad Retirement"
				If stat_unea_five_type_code(each_memb) = "17" Then stat_unea_five_type_info(each_memb) = "Other Retirement"
				If stat_unea_five_type_code(each_memb) = "18" Then stat_unea_five_type_info(each_memb) = "Military Entitlement"
				If stat_unea_five_type_code(each_memb) = "19" Then stat_unea_five_type_info(each_memb) = "Foster Care Child Requesting SNAP"
				If stat_unea_five_type_code(each_memb) = "20" Then stat_unea_five_type_info(each_memb) = "Foster Care Child NOT Requesting SNAP"
				If stat_unea_five_type_code(each_memb) = "21" Then stat_unea_five_type_info(each_memb) = "Foster Care Adult Requesting SNAP"
				If stat_unea_five_type_code(each_memb) = "22" Then stat_unea_five_type_info(each_memb) = "Foster Care Adult NOT Requesting SNAP"
				If stat_unea_five_type_code(each_memb) = "23" Then stat_unea_five_type_info(each_memb) = "Dividends"
				If stat_unea_five_type_code(each_memb) = "24" Then stat_unea_five_type_info(each_memb) = "Interest"
				If stat_unea_five_type_code(each_memb) = "25" Then stat_unea_five_type_info(each_memb) = "Counted Gifts or Prizes"
				If stat_unea_five_type_code(each_memb) = "26" Then stat_unea_five_type_info(each_memb) = "Strike Benefit"
				If stat_unea_five_type_code(each_memb) = "27" Then stat_unea_five_type_info(each_memb) = "Contract for Deed"
				If stat_unea_five_type_code(each_memb) = "28" Then stat_unea_five_type_info(each_memb) = "Illegal Income"
				If stat_unea_five_type_code(each_memb) = "29" Then stat_unea_five_type_info(each_memb) = "Other Countable"
				If stat_unea_five_type_code(each_memb) = "30" Then stat_unea_five_type_info(each_memb) = "Infrequent, <30, Not Counted"
				If stat_unea_five_type_code(each_memb) = "31" Then stat_unea_five_type_info(each_memb) = "Other SNAP Only"
				If stat_unea_five_type_code(each_memb) = "08" Then stat_unea_five_type_info(each_memb) = "Direct Child Support"
				If stat_unea_five_type_code(each_memb) = "35" Then stat_unea_five_type_info(each_memb) = "Direct Spousal Support"
				If stat_unea_five_type_code(each_memb) = "36" Then stat_unea_five_type_info(each_memb) = "Disbursed Child Support"
				If stat_unea_five_type_code(each_memb) = "37" Then stat_unea_five_type_info(each_memb) = "Disbursed Spousal Support"
				If stat_unea_five_type_code(each_memb) = "39" Then stat_unea_five_type_info(each_memb) = "Disbursed Child Support Arrears"
				If stat_unea_five_type_code(each_memb) = "40" Then stat_unea_five_type_info(each_memb) = "Disbursed Spousal Support Arrears"
				If stat_unea_five_type_code(each_memb) = "43" Then stat_unea_five_type_info(each_memb) = "Disbursed Excess Child Support"
				If stat_unea_five_type_code(each_memb) = "44" Then stat_unea_five_type_info(each_memb) = "MSA - Excess Income for SSI"
				If stat_unea_five_type_code(each_memb) = "45" Then stat_unea_five_type_info(each_memb) = "County 88 Child Support"
				If stat_unea_five_type_code(each_memb) = "46" Then stat_unea_five_type_info(each_memb) = "County 88 Gaming"
				If stat_unea_five_type_code(each_memb) = "47" Then stat_unea_five_type_info(each_memb) = "Counted Tribal Income"
				If stat_unea_five_type_code(each_memb) = "48" Then stat_unea_five_type_info(each_memb) = "Trust income"
				If stat_unea_five_type_code(each_memb) = "49" Then stat_unea_five_type_info(each_memb) = "Non-Recurring Income > $60 per Quarter"
				EMReadScreen stat_unea_five_verif_code(each_memb), 1, 5, 65
				If stat_unea_five_verif_code(each_memb) = "1" Then stat_unea_five_verif_info(each_memb) = "Copy of Checks"
				If stat_unea_five_verif_code(each_memb) = "2" Then stat_unea_five_verif_info(each_memb) = "Award Letter"
				If stat_unea_five_verif_code(each_memb) = "3" Then stat_unea_five_verif_info(each_memb) = "System Initiated Verif"
				If stat_unea_five_verif_code(each_memb) = "4" Then stat_unea_five_verif_info(each_memb) = "Collateral Statement"
				If stat_unea_five_verif_code(each_memb) = "5" Then stat_unea_five_verif_info(each_memb) = "Pending Out of State Verif"
				If stat_unea_five_verif_code(each_memb) = "6" Then stat_unea_five_verif_info(each_memb) = "Other Document"
				If stat_unea_five_verif_code(each_memb) = "7" Then stat_unea_five_verif_info(each_memb) = "Worker Initiated Verif"
				If stat_unea_five_verif_code(each_memb) = "8" Then stat_unea_five_verif_info(each_memb) = "RI Stubs"
				If stat_unea_five_verif_code(each_memb) = "N" Then stat_unea_five_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_unea_five_inc_start_date(each_memb), 8, 7, 37
				EMReadScreen stat_unea_five_inc_end_date(each_memb), 8, 7, 68

				stat_unea_five_inc_start_date(each_memb) = replace(stat_unea_five_inc_start_date(each_memb), " ", "/")
				stat_unea_five_inc_end_date(each_memb) = replace(stat_unea_five_inc_end_date(each_memb), " ", "/")
				iF stat_unea_five_inc_end_date(each_memb) = "__/__/__" Then stat_unea_five_inc_end_date(each_memb) = ""

				Call write_value_and_transmit("X", 10, 26)
				EMReadScreen stat_unea_five_snap_pic_pay_freq(each_memb), 1, 5, 64
				EMReadScreen stat_unea_five_snap_pic_ave_inc_per_pay(each_memb), 10, 17, 52
				EMReadScreen stat_unea_five_snap_pic_prosp_monthly_inc(each_memb), 10, 18, 52

				If stat_unea_five_snap_pic_pay_freq(each_memb) = "_" Then stat_unea_five_snap_pic_pay_freq(each_memb) = ""
				If stat_unea_five_snap_pic_pay_freq(each_memb) = "1" Then stat_unea_five_snap_pic_pay_freq(each_memb) = "Monthly"
				If stat_unea_five_snap_pic_pay_freq(each_memb) = "2" Then stat_unea_five_snap_pic_pay_freq(each_memb) = "Semi-Monthly"
				If stat_unea_five_snap_pic_pay_freq(each_memb) = "3" Then stat_unea_five_snap_pic_pay_freq(each_memb) = "BiWeekly"
				If stat_unea_five_snap_pic_pay_freq(each_memb) = "4" Then stat_unea_five_snap_pic_pay_freq(each_memb) = "Weekly"
				stat_unea_five_snap_pic_ave_inc_per_pay(each_memb) = trim(stat_unea_five_snap_pic_ave_inc_per_pay(each_memb))
				stat_unea_five_snap_pic_prosp_monthly_inc(each_memb) = trim(stat_unea_five_snap_pic_prosp_monthly_inc(each_memb))
				PF3

			End If

		Next

		call navigate_to_MAXIS_screen("STAT", "ACCT")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_acct_one_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_acct_one_exists(each_memb) = False

			If stat_acct_one_exists(each_memb) = True Then
				EMReadScreen stat_acct_one_type(each_memb), 2, 6, 44
				EMReadScreen stat_acct_one_balence(each_memb), 8, 10, 46
				EMReadScreen stat_acct_one_count_snap_yn(each_memb), 1, 14, 57
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "02", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_acct_two_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_acct_two_exists(each_memb) = False

			If stat_acct_two_exists(each_memb) = True Then
				EMReadScreen stat_acct_two_type(each_memb), 2, 6, 44
				EMReadScreen stat_acct_two_balence(each_memb), 8, 10, 46
				EMReadScreen stat_acct_two_count_snap_yn(each_memb), 1, 14, 57
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "03", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_acct_three_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_acct_three_exists(each_memb) = False

			If stat_acct_three_exists(each_memb) = True Then
				EMReadScreen stat_acct_three_type(each_memb), 2, 6, 44
				EMReadScreen stat_acct_three_balence(each_memb), 8, 10, 46
				EMReadScreen stat_acct_three_count_snap_yn(each_memb), 1, 14, 57
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "04", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_acct_four_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_acct_four_exists(each_memb) = False

			If stat_acct_four_exists(each_memb) = True Then
				EMReadScreen stat_acct_four_type(each_memb), 2, 6, 44
				EMReadScreen stat_acct_four_balence(each_memb), 8, 10, 46
				EMReadScreen stat_acct_four_count_snap_yn(each_memb), 1, 14, 57
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "05", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_acct_five_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_acct_five_exists(each_memb) = False

			If stat_acct_five_exists(each_memb) = True Then
				EMReadScreen stat_acct_five_type(each_memb), 2, 6, 44
				EMReadScreen stat_acct_five_balence(each_memb), 8, 10, 46
				EMReadScreen stat_acct_five_count_snap_yn(each_memb), 1, 14, 57
			End If
		Next

		call navigate_to_MAXIS_screen("STAT", "SHEL")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			Call write_value_and_transmit(stat_memb_ref_numb(each_memb), 20, 76)

			stat_shel_exists(each_memb) = False
			EMReadScreen shel_version, 1, 2, 73
			If shel_version = "1" Then
				stat_shel_exists(each_memb) = True

				EMReadScreen stat_shel_subsidized_yn(each_memb), 1, 6, 46
				EMReadScreen stat_shel_shared_yn(each_memb), 1, 6, 64
				EMReadScreen stat_shel_paid_to(each_memb), 25, 7, 50

				EMReadScreen stat_shel_retro_rent_amount(each_memb), 8, 11, 37
				EMReadScreen stat_shel_retro_rent_verif_code(each_memb), 2, 11, 48
				If stat_shel_retro_rent_verif_code(each_memb) = "__" Then stat_shel_retro_rent_verif_info(each_memb) = ""
				If stat_shel_retro_rent_verif_code(each_memb) = "SF" Then stat_shel_retro_rent_verif_info(each_memb) = "Shelter Form"
				If stat_shel_retro_rent_verif_code(each_memb) = "LE" Then stat_shel_retro_rent_verif_info(each_memb) = "Lease"
				If stat_shel_retro_rent_verif_code(each_memb) = "RE" Then stat_shel_retro_rent_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_retro_rent_verif_code(each_memb) = "OT" Then stat_shel_retro_rent_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_rent_verif_code(each_memb) = "NC" Then stat_shel_retro_rent_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_retro_rent_verif_code(each_memb) = "PC" Then stat_shel_retro_rent_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_retro_rent_verif_code(each_memb) = "NO" Then stat_shel_retro_rent_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_shel_prosp_rent_amount(each_memb), 8, 11, 56
				EMReadScreen stat_shel_prosp_rent_verif_code(each_memb), 2, 11, 67
				If stat_shel_prosp_rent_verif_code(each_memb) = "__" Then stat_shel_prosp_rent_verif_info(each_memb) = ""
				If stat_shel_prosp_rent_verif_code(each_memb) = "SF" Then stat_shel_prosp_rent_verif_info(each_memb) = "Shelter Form"
				If stat_shel_prosp_rent_verif_code(each_memb) = "LE" Then stat_shel_prosp_rent_verif_info(each_memb) = "Lease"
				If stat_shel_prosp_rent_verif_code(each_memb) = "RE" Then stat_shel_prosp_rent_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_prosp_rent_verif_code(each_memb) = "OT" Then stat_shel_prosp_rent_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_rent_verif_code(each_memb) = "NC" Then stat_shel_prosp_rent_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_rent_verif_code(each_memb) = "PC" Then stat_shel_prosp_rent_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_rent_verif_code(each_memb) = "NO" Then stat_shel_prosp_rent_verif_info(each_memb) = "No Verif Provided"

				EMReadScreen stat_shel_retro_lot_rent_amount(each_memb), 8, 12, 37
				EMReadScreen stat_shel_retro_lot_rent_verif_code(each_memb), 2, 12, 48
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "__" Then stat_shel_retro_lot_rent_verif_info(each_memb) = ""
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "LE" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "Lease"
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "RE" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "BI" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "Billing Statement"
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "OT" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "NC" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "PC" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_retro_lot_rent_verif_code(each_memb) = "NO" Then stat_shel_retro_lot_rent_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_shel_prosp_lot_rent_amount(each_memb), 8, 12, 56
				EMReadScreen stat_shel_prosp_lot_rent_verif_code(each_memb), 2, 12, 67
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "__" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = ""
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "LE" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "Lease"
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "RE" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "BI" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "Billing Statement"
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "OT" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "NC" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "PC" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_lot_rent_verif_code(each_memb) = "NO" Then stat_shel_prosp_lot_rent_verif_info(each_memb) = "No Verif Provided"

				EMReadScreen stat_shel_retro_mortgage_amount(each_memb), 8, 13, 37
				EMReadScreen stat_shel_retro_mortgage_verif_code(each_memb), 2, 13, 48
				If stat_shel_retro_mortgage_verif_code(each_memb) = "__" Then stat_shel_retro_mortgage_verif_info(each_memb) = ""
				If stat_shel_retro_mortgage_verif_code(each_memb) = "MO" Then stat_shel_retro_mortgage_verif_info(each_memb) = "Mortgage Payment"
				If stat_shel_retro_mortgage_verif_code(each_memb) = "CD" Then stat_shel_retro_mortgage_verif_info(each_memb) = "Contract for Deed"
				If stat_shel_retro_mortgage_verif_code(each_memb) = "OT" Then stat_shel_retro_mortgage_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_mortgage_verif_code(each_memb) = "NC" Then stat_shel_retro_mortgage_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_retro_mortgage_verif_code(each_memb) = "PC" Then stat_shel_retro_mortgage_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_retro_mortgage_verif_code(each_memb) = "NO" Then stat_shel_retro_mortgage_verif_info(each_memb) = "No Verif provided"
				EMReadScreen stat_shel_prosp_mortgage_amount(each_memb), 8, 13, 56
				EMReadScreen stat_shel_prosp_mortgage_verif_code(each_memb), 2, 13, 67
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "__" Then stat_shel_prosp_mortgage_verif_info(each_memb) = ""
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "MO" Then stat_shel_prosp_mortgage_verif_info(each_memb) = "Mortgage Payment"
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "CD" Then stat_shel_prosp_mortgage_verif_info(each_memb) = "Contract for Deed"
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "OT" Then stat_shel_prosp_mortgage_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "NC" Then stat_shel_prosp_mortgage_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "PC" Then stat_shel_prosp_mortgage_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_mortgage_verif_code(each_memb) = "NO" Then stat_shel_prosp_mortgage_verif_info(each_memb) = "No Verif provided"

				EMReadScreen stat_shel_retro_insurance_amount(each_memb), 8, 14, 37
				EMReadScreen stat_shel_retro_insurance_verif_code(each_memb), 2, 14, 48
				If stat_shel_retro_insurance_verif_code(each_memb) = "__" Then stat_shel_retro_insurance_verif_info(each_memb) = ""
				If stat_shel_retro_insurance_verif_code(each_memb) = "BI" Then stat_shel_retro_insurance_verif_info(each_memb) = "Billing Statement"
				If stat_shel_retro_insurance_verif_code(each_memb) = "OT" Then stat_shel_retro_insurance_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_insurance_verif_code(each_memb) = "NC" Then stat_shel_retro_insurance_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_retro_insurance_verif_code(each_memb) = "PC" Then stat_shel_retro_insurance_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_retro_insurance_verif_code(each_memb) = "NO" Then stat_shel_retro_insurance_verif_info(each_memb) = "No Verif provided"
				EMReadScreen stat_shel_prosp_insurance_amount(each_memb), 8, 14, 56
				EMReadScreen stat_shel_prosp_insurance_verif_code(each_memb), 2, 14, 67
				If stat_shel_prosp_insurance_verif_code(each_memb) = "__" Then stat_shel_prosp_insurance_verif_info(each_memb) = ""
				If stat_shel_prosp_insurance_verif_code(each_memb) = "BI" Then stat_shel_prosp_insurance_verif_info(each_memb) = "Billing Statement"
				If stat_shel_prosp_insurance_verif_code(each_memb) = "OT" Then stat_shel_prosp_insurance_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_insurance_verif_code(each_memb) = "NC" Then stat_shel_prosp_insurance_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_insurance_verif_code(each_memb) = "PC" Then stat_shel_prosp_insurance_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_insurance_verif_code(each_memb) = "NO" Then stat_shel_prosp_insurance_verif_info(each_memb) = "No Verif provided"

				EMReadScreen stat_shel_retro_taxes_amount(each_memb), 8, 15, 37
				EMReadScreen stat_shel_retro_taxes_verif_code(each_memb), 2, 15, 48
				If stat_shel_retro_taxes_verif_code(each_memb) = "__" Then stat_shel_retro_taxes_verif_info(each_memb) = ""
				If stat_shel_retro_taxes_verif_code(each_memb) = "TX" Then stat_shel_retro_taxes_verif_info(each_memb) = "Property Tax Statement"
				If stat_shel_retro_taxes_verif_code(each_memb) = "OT" Then stat_shel_retro_taxes_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_taxes_verif_code(each_memb) = "NC" Then stat_shel_retro_taxes_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_retro_taxes_verif_code(each_memb) = "PC" Then stat_shel_retro_taxes_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_retro_taxes_verif_code(each_memb) = "NO" Then stat_shel_retro_taxes_verif_info(each_memb) = "No Verif provided"
				EMReadScreen stat_shel_prosp_taxes_amount(each_memb), 8, 15, 56
				EMReadScreen stat_shel_prosp_taxes_verif_code(each_memb), 2, 15, 67
				If stat_shel_prosp_taxes_verif_code(each_memb) = "__" Then stat_shel_prosp_taxes_verif_info(each_memb) = ""
				If stat_shel_prosp_taxes_verif_code(each_memb) = "TX" Then stat_shel_prosp_taxes_verif_info(each_memb) = "Property Tax Statement"
				If stat_shel_prosp_taxes_verif_code(each_memb) = "OT" Then stat_shel_prosp_taxes_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_taxes_verif_code(each_memb) = "NC" Then stat_shel_prosp_taxes_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_taxes_verif_code(each_memb) = "PC" Then stat_shel_prosp_taxes_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_taxes_verif_code(each_memb) = "NO" Then stat_shel_prosp_taxes_verif_info(each_memb) = "No Verif provided"

				EMReadScreen stat_shel_retro_room_amount(each_memb), 8, 16, 37
				EMReadScreen stat_shel_retro_room_verif_code(each_memb), 2, 16, 48
				If stat_shel_retro_room_verif_code(each_memb) = "__" Then stat_shel_retro_room_verif_info(each_memb) = ""
				If stat_shel_retro_room_verif_code(each_memb) = "SF" Then stat_shel_prosp_rent_verif_info(each_memb) = "Shelter Form"
				If stat_shel_retro_room_verif_code(each_memb) = "LE" Then stat_shel_prosp_rent_verif_info(each_memb) = "Lease"
				If stat_shel_prosp_rent_verif_code(each_memb) = "RE" Then stat_shel_prosp_rent_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_prosp_rent_verif_code(each_memb) = "OT" Then stat_shel_prosp_rent_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_rent_verif_code(each_memb) = "NC" Then stat_shel_prosp_rent_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_rent_verif_code(each_memb) = "PC" Then stat_shel_prosp_rent_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_rent_verif_code(each_memb) = "NO" Then stat_shel_prosp_rent_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_shel_prosp_room_amount(each_memb), 8, 16, 56
				EMReadScreen stat_shel_prosp_room_verif_code(each_memb), 2, 16, 67
				If stat_shel_prosp_room_verif_code(each_memb) = "__" Then stat_shel_prosp_room_verif_info(each_memb) = ""
				If stat_shel_prosp_room_verif_code(each_memb) = "SF" Then stat_shel_prosp_room_verif_info(each_memb) = "Shelter Form"
				If stat_shel_prosp_room_verif_code(each_memb) = "LE" Then stat_shel_prosp_room_verif_info(each_memb) = "Lease"
				If stat_shel_prosp_room_verif_code(each_memb) = "RE" Then stat_shel_prosp_room_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_prosp_room_verif_code(each_memb) = "OT" Then stat_shel_prosp_room_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_room_verif_code(each_memb) = "NC" Then stat_shel_prosp_room_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_room_verif_code(each_memb) = "PC" Then stat_shel_prosp_room_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_room_verif_code(each_memb) = "NO" Then stat_shel_prosp_room_verif_info(each_memb) = "No Verif Provided"

				EMReadScreen stat_shel_retro_garage_amount(each_memb), 8, 17, 37
				EMReadScreen stat_shel_retro_garage_verif_code(each_memb), 2, 17, 48
				If stat_shel_retro_garage_verif_code(each_memb) = "__" Then stat_shel_retro_garage_verif_info(each_memb) = ""
				If stat_shel_retro_garage_verif_code(each_memb) = "SF" Then stat_shel_retro_garage_verif_info(each_memb) = "Shelter Form"
				If stat_shel_retro_garage_verif_code(each_memb) = "LE" Then stat_shel_retro_garage_verif_info(each_memb) = "Lease"
				If stat_shel_retro_garage_verif_code(each_memb) = "RE" Then stat_shel_retro_garage_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_retro_garage_verif_code(each_memb) = "OT" Then stat_shel_retro_garage_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_garage_verif_code(each_memb) = "NC" Then stat_shel_retro_garage_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_retro_garage_verif_code(each_memb) = "PC" Then stat_shel_retro_garage_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_retro_garage_verif_code(each_memb) = "NO" Then stat_shel_retro_garage_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_shel_prosp_garage_amount(each_memb), 8, 17, 56
				EMReadScreen stat_shel_prosp_garage_verif_code(each_memb), 2, 17, 67
				If stat_shel_prosp_garage_verif_code(each_memb) = "__" Then stat_shel_prosp_garage_verif_info(each_memb) = ""
				If stat_shel_prosp_garage_verif_code(each_memb) = "SF" Then stat_shel_prosp_garage_verif_info(each_memb) = "Shelter Form"
				If stat_shel_prosp_garage_verif_code(each_memb) = "LE" Then stat_shel_prosp_garage_verif_info(each_memb) = "Lease"
				If stat_shel_prosp_garage_verif_code(each_memb) = "RE" Then stat_shel_prosp_garage_verif_info(each_memb) = "Rent Receipts"
				If stat_shel_prosp_garage_verif_code(each_memb) = "OT" Then stat_shel_prosp_garage_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_garage_verif_code(each_memb) = "NC" Then stat_shel_prosp_garage_verif_info(each_memb) = "Change Reported, No Verif, Negative Impact"
				If stat_shel_prosp_garage_verif_code(each_memb) = "PC" Then stat_shel_prosp_garage_verif_info(each_memb) = "Change Reported, No Verif, Positive Impact"
				If stat_shel_prosp_garage_verif_code(each_memb) = "NO" Then stat_shel_prosp_garage_verif_info(each_memb) = "No Verif Provided"

				EMReadScreen stat_shel_retro_subsidy_amount(each_memb), 8, 18, 37
				EMReadScreen stat_shel_retro_subsidy_verif_code(each_memb), 2, 18, 48
				If stat_shel_retro_subsidy_verif_code(each_memb) = "__" Then stat_shel_retro_subsidy_verif_info(each_memb) = ""
				If stat_shel_retro_subsidy_verif_code(each_memb) = "SF" Then stat_shel_retro_subsidy_verif_info(each_memb) = "Shelter Form"
				If stat_shel_retro_subsidy_verif_code(each_memb) = "LE" Then stat_shel_retro_subsidy_verif_info(each_memb) = "Lease"
				If stat_shel_retro_subsidy_verif_code(each_memb) = "OT" Then stat_shel_retro_subsidy_verif_info(each_memb) = "Other Document"
				If stat_shel_retro_subsidy_verif_code(each_memb) = "NO" Then stat_shel_retro_subsidy_verif_info(each_memb) = "No Verif Provided"
				EMReadScreen stat_shel_prosp_subsidy_amount(each_memb), 8, 18, 56
				EMReadScreen stat_shel_prosp_subsidy_verif_code(each_memb), 2, 18, 67
				If stat_shel_prosp_subsidy_verif_code(each_memb) = "__" Then stat_shel_prosp_subsidy_verif_info(each_memb) = ""
				If stat_shel_prosp_subsidy_verif_code(each_memb) = "SF" Then stat_shel_prosp_subsidy_verif_info(each_memb) = "Shelter Form"
				If stat_shel_prosp_subsidy_verif_code(each_memb) = "LE" Then stat_shel_prosp_subsidy_verif_info(each_memb) = "Lease"
				If stat_shel_prosp_subsidy_verif_code(each_memb) = "OT" Then stat_shel_prosp_subsidy_verif_info(each_memb) = "Other Document"
				If stat_shel_prosp_subsidy_verif_code(each_memb) = "NO" Then stat_shel_prosp_subsidy_verif_info(each_memb) = "No Verif Provided"



				stat_shel_prosp_rent_amount(each_memb) = trim(stat_shel_prosp_rent_amount(each_memb))
				If stat_shel_prosp_rent_amount(each_memb) = "________" Then stat_shel_prosp_rent_amount(each_memb) = 0
				stat_shel_prosp_rent_amount(each_memb) = stat_shel_prosp_rent_amount(each_memb)*1

				stat_shel_prosp_lot_rent_amount(each_memb) = trim(stat_shel_prosp_lot_rent_amount(each_memb))
				If stat_shel_prosp_lot_rent_amount(each_memb) = "________" Then stat_shel_prosp_lot_rent_amount(each_memb) = 0
				stat_shel_prosp_lot_rent_amount(each_memb) = stat_shel_prosp_lot_rent_amount(each_memb)*1

				stat_shel_prosp_mortgage_amount(each_memb) = trim(stat_shel_prosp_mortgage_amount(each_memb))
				If stat_shel_prosp_mortgage_amount(each_memb) = "________" Then stat_shel_prosp_mortgage_amount(each_memb) = 0
				stat_shel_prosp_mortgage_amount(each_memb) = stat_shel_prosp_mortgage_amount(each_memb)*1

				stat_shel_prosp_insurance_amount(each_memb) = trim(stat_shel_prosp_insurance_amount(each_memb))
				If stat_shel_prosp_insurance_amount(each_memb) = "________" Then stat_shel_prosp_insurance_amount(each_memb) = 0
				stat_shel_prosp_insurance_amount(each_memb) = stat_shel_prosp_insurance_amount(each_memb)*1

				stat_shel_prosp_taxes_amount(each_memb) = trim(stat_shel_prosp_taxes_amount(each_memb))
				If stat_shel_prosp_taxes_amount(each_memb) = "________" Then stat_shel_prosp_taxes_amount(each_memb) = 0
				stat_shel_prosp_taxes_amount(each_memb) = stat_shel_prosp_taxes_amount(each_memb)*1

				stat_shel_prosp_room_amount(each_memb) = trim(stat_shel_prosp_room_amount(each_memb))
				If stat_shel_prosp_room_amount(each_memb) = "________" Then stat_shel_prosp_room_amount(each_memb) = 0
				stat_shel_prosp_room_amount(each_memb) = stat_shel_prosp_room_amount(each_memb)*1

				stat_shel_prosp_garage_amount(each_memb) = trim(stat_shel_prosp_garage_amount(each_memb))
				If stat_shel_prosp_garage_amount(each_memb) = "________" Then stat_shel_prosp_garage_amount(each_memb) = 0
				stat_shel_prosp_garage_amount(each_memb) = stat_shel_prosp_garage_amount(each_memb)*1

				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_rent_amount(each_memb)
				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_lot_rent_amount(each_memb)
				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_mortgage_amount(each_memb)
				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_insurance_amount(each_memb)
				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_taxes_amount(each_memb)
				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_room_amount(each_memb)
				stat_shel_prosp_all_total = stat_shel_prosp_all_total + stat_shel_prosp_garage_amount(each_memb)



			End If
		Next


		call navigate_to_MAXIS_screen("STAT", "DISQ")
		For each_memb = 0 to UBound(stat_memb_ref_numb)
			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "01", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_disq_one_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_disq_one_exists(each_memb) = False

			If stat_disq_one_exists(each_memb) = True Then
				EMReadScreen stat_disq_one_program(each_memb), 2, 6, 54
				If stat_disq_one_program(each_memb) = "__" Then stat_disq_one_program(each_memb) = ""
				If stat_disq_one_program(each_memb) = "AF" Then stat_disq_one_program(each_memb) = "AFDC"
				If stat_disq_one_program(each_memb) = "CC" Then stat_disq_one_program(each_memb) = "Child Care Assistance"
				If stat_disq_one_program(each_memb) = "DW" Then stat_disq_one_program(each_memb) = "DWP"
				If stat_disq_one_program(each_memb) = "FG" Then stat_disq_one_program(each_memb) = "Family General Assistance"
				If stat_disq_one_program(each_memb) = "FS" Then stat_disq_one_program(each_memb) = "SNAP"
				If stat_disq_one_program(each_memb) = "GA" Then stat_disq_one_program(each_memb) = "General Asssistance"
				If stat_disq_one_program(each_memb) = "GR" Then stat_disq_one_program(each_memb) = "GRH"
				If stat_disq_one_program(each_memb) = "IM" Then stat_disq_one_program(each_memb) = "IMD"
				If stat_disq_one_program(each_memb) = "MA" Then stat_disq_one_program(each_memb) = "Medical Assistance"
				If stat_disq_one_program(each_memb) = "MF" Then stat_disq_one_program(each_memb) = "MFIP"
				If stat_disq_one_program(each_memb) = "MS" Then stat_disq_one_program(each_memb) = "MN Supplemental Aid"
				If stat_disq_one_program(each_memb) = "QI" Then stat_disq_one_program(each_memb) = "QI-1"
				If stat_disq_one_program(each_memb) = "QM" Then stat_disq_one_program(each_memb) = "QMB"
				If stat_disq_one_program(each_memb) = "QW" Then stat_disq_one_program(each_memb) = "QWD"
				If stat_disq_one_program(each_memb) = "RC" Then stat_disq_one_program(each_memb) = "Refugee Cash Assistance"
				If stat_disq_one_program(each_memb) = "RM" Then stat_disq_one_program(each_memb) = "RMA"
				If stat_disq_one_program(each_memb) = "SL" Then stat_disq_one_program(each_memb) = "SLMB"
				If stat_disq_one_program(each_memb) = "WB" Then stat_disq_one_program(each_memb) = "Work Benefit Program"
				If stat_disq_one_program(each_memb) = "4E" Then stat_disq_one_program(each_memb) = "Title IV-E Foster Care"
				EMReadScreen stat_disq_one_type_code(each_memb), 2, 6, 64
				If stat_disq_one_type_code(each_memb) = "__" Then stat_disq_one_type_info(each_memb) = ""
				If stat_disq_one_type_code(each_memb) = "02" Then stat_disq_one_type_info(each_memb) = "SNAP Fraud - 1st Disq"
				If stat_disq_one_type_code(each_memb) = "03" Then stat_disq_one_type_info(each_memb) = "SNAP Fraud - 2md Disq"
				If stat_disq_one_type_code(each_memb) = "04" Then stat_disq_one_type_info(each_memb) = "SNAP Fraud - 3rd Disq"
				If stat_disq_one_type_code(each_memb) = "06" Then stat_disq_one_type_info(each_memb) = "Non-Coop with State QC"
				If stat_disq_one_type_code(each_memb) = "07" Then stat_disq_one_type_info(each_memb) = "Non-Coop with Federal QC"
				If stat_disq_one_type_code(each_memb) = "08" Then stat_disq_one_type_info(each_memb) = "RCA Non-Comply with E&T"
				If stat_disq_one_type_code(each_memb) = "11" Then stat_disq_one_type_info(each_memb) = "Voluntary Quit"
				If stat_disq_one_type_code(each_memb) = "12" Then stat_disq_one_type_info(each_memb) = "Improper Transfer of Assets"
				If stat_disq_one_type_code(each_memb) = "13" Then stat_disq_one_type_info(each_memb) = "Lump Sum"
				If stat_disq_one_type_code(each_memb) = "14" Then stat_disq_one_type_info(each_memb) = "IEVS Non-Coop"
				If stat_disq_one_type_code(each_memb) = "15" Then stat_disq_one_type_info(each_memb) = "Cash Fraud - Time Set by Court"
				If stat_disq_one_type_code(each_memb) = "16" Then stat_disq_one_type_info(each_memb) = "Cash Fraud - 1st Disq"
				If stat_disq_one_type_code(each_memb) = "17" Then stat_disq_one_type_info(each_memb) = "Cash Fraud - 2nd Disq"
				If stat_disq_one_type_code(each_memb) = "18" Then stat_disq_one_type_info(each_memb) = "Cash Fraud - 3rd Disq"
				If stat_disq_one_type_code(each_memb) = "20" Then stat_disq_one_type_info(each_memb) = "Improper Transfer of Income"
				If stat_disq_one_type_code(each_memb) = "23" Then stat_disq_one_type_info(each_memb) = "Fleeing Felon, Violating Parole/Probation, Explosives"
				If stat_disq_one_type_code(each_memb) = "26" Then stat_disq_one_type_info(each_memb) = "Family Cash Falsify Residence for Duplicate Assistance"
				If stat_disq_one_type_code(each_memb) = "28" Then stat_disq_one_type_info(each_memb) = "Convicted of Drug Felony - Failed Drug Test"
				If stat_disq_one_type_code(each_memb) = "29" Then stat_disq_one_type_info(each_memb) = "US Citizenship/ID Verif Non-Coop"
				If stat_disq_one_type_code(each_memb) = "30" Then stat_disq_one_type_info(each_memb) = "Immigration Status Verif Non-Coop"
				If stat_disq_one_type_code(each_memb) = "31" Then stat_disq_one_type_info(each_memb) = "EBT Misuse - 1st Disq"
				If stat_disq_one_type_code(each_memb) = "32" Then stat_disq_one_type_info(each_memb) = "EBT Misuse - 2nd Disq"
				If stat_disq_one_type_code(each_memb) = "33" Then stat_disq_one_type_info(each_memb) = "EBT Misuse - 3rd Disq"
				EMReadScreen stat_disq_one_begin_date(each_memb), 8, 8, 64
				stat_disq_one_begin_date(each_memb) = replace(stat_disq_one_begin_date(each_memb), " ", "/")
				If stat_disq_one_begin_date(each_memb) = "__/__/__" Then stat_disq_one_begin_date(each_memb) = ""
				EMReadScreen stat_disq_one_end_date(each_memb), 8, 9, 64
				stat_disq_one_end_date(each_memb) = replace(stat_disq_one_end_date(each_memb), " ", "/")
				If stat_disq_one_end_date(each_memb) = "__/__/__" Then stat_disq_one_end_date(each_memb) = ""
				EMReadScreen stat_disq_one_cure_reason_code(each_memb), 1, 11, 64
				If stat_disq_one_cure_reason_code(each_memb) = "_" Then stat_disq_one_cure_reason_info(each_memb) = ""
				If stat_disq_one_cure_reason_code(each_memb) = "A" Then stat_disq_one_cure_reason_info(each_memb) = "No longer Fleeing Felon/Parole Violation"
				If stat_disq_one_cure_reason_code(each_memb) = "0" Then stat_disq_one_cure_reason_info(each_memb) = "Property Returned/Adequate Compensation"
				If stat_disq_one_cure_reason_code(each_memb) = "1" Then stat_disq_one_cure_reason_info(each_memb) = "Return to Same Job"
				If stat_disq_one_cure_reason_code(each_memb) = "2" Then stat_disq_one_cure_reason_info(each_memb) = "Accept Comparable Employment"
				If stat_disq_one_cure_reason_code(each_memb) = "6" Then stat_disq_one_cure_reason_info(each_memb) = "Cooperate with State QC"
				If stat_disq_one_cure_reason_code(each_memb) = "7" Then stat_disq_one_cure_reason_info(each_memb) = "Cooperate with Federal QC"
				If stat_disq_one_cure_reason_code(each_memb) = "8" Then stat_disq_one_cure_reason_info(each_memb) = "Lump Sum Recalculated"
				EMReadScreen stat_disq_one_fraud_determination_date(each_memb), 8, 13, 64
				stat_disq_one_fraud_determination_date(each_memb) = replace(stat_disq_one_fraud_determination_date(each_memb), " ", "/")
				If stat_disq_one_fraud_determination_date(each_memb) = "__/__/__" Then stat_disq_one_fraud_determination_date(each_memb) = ""
				EMReadScreen stat_disq_one_county_of_fraud(each_memb), 2, 15, 64
				EMReadScreen stat_disq_one_state_of_fraud(each_memb), 2, 16, 64
				EMReadScreen stat_disq_one_SNAP_trafficking_yn(each_memb), 1, 17, 64
				EMReadScreen stat_disq_one_SNAP_offense_code(each_memb), 2, 18, 64
				If stat_disq_one_SNAP_offense_code(each_memb) = "__" Then stat_disq_one_SNAP_offense_info(each_memb) = ""
				If stat_disq_one_SNAP_offense_code(each_memb) = "AL" Then stat_disq_one_SNAP_offense_info(each_memb) = "Alcohol"
				If stat_disq_one_SNAP_offense_code(each_memb) = "DR" Then stat_disq_one_SNAP_offense_info(each_memb) = "Drugs"
				If stat_disq_one_SNAP_offense_code(each_memb) = "GU" Then stat_disq_one_SNAP_offense_info(each_memb) = "Guns"
				If stat_disq_one_SNAP_offense_code(each_memb) = "OT" Then stat_disq_one_SNAP_offense_info(each_memb) = "Other"

				If stat_disq_one_type_code(each_memb) = "02" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "03" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "04" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "15" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "16" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "17" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "18" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "23" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "26" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "28" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "31" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "32" Then stat_disq_one_source(each_memb) = "DISQUAL"
				If stat_disq_one_type_code(each_memb) = "32" Then stat_disq_one_source(each_memb) = "DISQUAL"

				If stat_disq_one_type_code(each_memb) = "06" Then stat_disq_one_source(each_memb) = "NON-COOP"
				If stat_disq_one_type_code(each_memb) = "07" Then stat_disq_one_source(each_memb) = "NON-COOP"
				If stat_disq_one_type_code(each_memb) = "08" Then stat_disq_one_source(each_memb) = "NON-COOP"
				If stat_disq_one_type_code(each_memb) = "14" Then stat_disq_one_source(each_memb) = "NON-COOP"
				If stat_disq_one_type_code(each_memb) = "29" Then stat_disq_one_source(each_memb) = "NON-COOP"
				If stat_disq_one_type_code(each_memb) = "30" Then stat_disq_one_source(each_memb) = "NON-COOP"

				If stat_disq_one_type_code(each_memb) = "12" Then stat_disq_one_source(each_memb) = "TRANSFER"
				If stat_disq_one_type_code(each_memb) = "20" Then stat_disq_one_source(each_memb) = "TRANSFER"

				If stat_disq_one_type_code(each_memb) = "11" Then stat_disq_one_source(each_memb) = "VOL QUIT"

				stat_disq_one_active(each_memb) = True
				If IsDate(stat_disq_one_end_date(each_memb)) = True Then
					If DateDiff("m", stat_disq_one_end_date(each_memb), current_month) >= 0 Then stat_disq_one_active(each_memb) = False
				End If
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "02", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_disq_two_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_disq_two_exists(each_memb) = False

			If stat_disq_two_exists(each_memb) = True Then
				EMReadScreen stat_disq_two_program(each_memb), 2, 6, 54
				If stat_disq_two_program(each_memb) = "__" Then stat_disq_two_program(each_memb) = ""
				If stat_disq_two_program(each_memb) = "AF" Then stat_disq_two_program(each_memb) = "AFDC"
				If stat_disq_two_program(each_memb) = "CC" Then stat_disq_two_program(each_memb) = "Child Care Assistance"
				If stat_disq_two_program(each_memb) = "DW" Then stat_disq_two_program(each_memb) = "DWP"
				If stat_disq_two_program(each_memb) = "FG" Then stat_disq_two_program(each_memb) = "Family General Assistance"
				If stat_disq_two_program(each_memb) = "FS" Then stat_disq_two_program(each_memb) = "SNAP"
				If stat_disq_two_program(each_memb) = "GA" Then stat_disq_two_program(each_memb) = "General Asssistance"
				If stat_disq_two_program(each_memb) = "GR" Then stat_disq_two_program(each_memb) = "GRH"
				If stat_disq_two_program(each_memb) = "IM" Then stat_disq_two_program(each_memb) = "IMD"
				If stat_disq_two_program(each_memb) = "MA" Then stat_disq_two_program(each_memb) = "Medical Assistance"
				If stat_disq_two_program(each_memb) = "MF" Then stat_disq_two_program(each_memb) = "MFIP"
				If stat_disq_two_program(each_memb) = "MS" Then stat_disq_two_program(each_memb) = "MN Supplemental Aid"
				If stat_disq_two_program(each_memb) = "QI" Then stat_disq_two_program(each_memb) = "QI-1"
				If stat_disq_two_program(each_memb) = "QM" Then stat_disq_two_program(each_memb) = "QMB"
				If stat_disq_two_program(each_memb) = "QW" Then stat_disq_two_program(each_memb) = "QWD"
				If stat_disq_two_program(each_memb) = "RC" Then stat_disq_two_program(each_memb) = "Refugee Cash Assistance"
				If stat_disq_two_program(each_memb) = "RM" Then stat_disq_two_program(each_memb) = "RMA"
				If stat_disq_two_program(each_memb) = "SL" Then stat_disq_two_program(each_memb) = "SLMB"
				If stat_disq_two_program(each_memb) = "WB" Then stat_disq_two_program(each_memb) = "Work Benefit Program"
				If stat_disq_two_program(each_memb) = "4E" Then stat_disq_two_program(each_memb) = "Title IV-E Foster Care"
				EMReadScreen stat_disq_two_type_code(each_memb), 2, 6, 64
				If stat_disq_two_type_code(each_memb) = "__" Then stat_disq_two_type_info(each_memb) = ""
				If stat_disq_two_type_code(each_memb) = "02" Then stat_disq_two_type_info(each_memb) = "SNAP Fraud - 1st Disq"
				If stat_disq_two_type_code(each_memb) = "03" Then stat_disq_two_type_info(each_memb) = "SNAP Fraud - 2md Disq"
				If stat_disq_two_type_code(each_memb) = "04" Then stat_disq_two_type_info(each_memb) = "SNAP Fraud - 3rd Disq"
				If stat_disq_two_type_code(each_memb) = "06" Then stat_disq_two_type_info(each_memb) = "Non-Coop with State QC"
				If stat_disq_two_type_code(each_memb) = "07" Then stat_disq_two_type_info(each_memb) = "Non-Coop with Federal QC"
				If stat_disq_two_type_code(each_memb) = "08" Then stat_disq_two_type_info(each_memb) = "RCA Non-Comply with E&T"
				If stat_disq_two_type_code(each_memb) = "11" Then stat_disq_two_type_info(each_memb) = "Voluntary Quit"
				If stat_disq_two_type_code(each_memb) = "12" Then stat_disq_two_type_info(each_memb) = "Improper Transfer of Assets"
				If stat_disq_two_type_code(each_memb) = "13" Then stat_disq_two_type_info(each_memb) = "Lump Sum"
				If stat_disq_two_type_code(each_memb) = "14" Then stat_disq_two_type_info(each_memb) = "IEVS Non-Coop"
				If stat_disq_two_type_code(each_memb) = "15" Then stat_disq_two_type_info(each_memb) = "Cash Fraud - Time Set by Court"
				If stat_disq_two_type_code(each_memb) = "16" Then stat_disq_two_type_info(each_memb) = "Cash Fraud - 1st Disq"
				If stat_disq_two_type_code(each_memb) = "17" Then stat_disq_two_type_info(each_memb) = "Cash Fraud - 2nd Disq"
				If stat_disq_two_type_code(each_memb) = "18" Then stat_disq_two_type_info(each_memb) = "Cash Fraud - 3rd Disq"
				If stat_disq_two_type_code(each_memb) = "20" Then stat_disq_two_type_info(each_memb) = "Improper Transfer of Income"
				If stat_disq_two_type_code(each_memb) = "23" Then stat_disq_two_type_info(each_memb) = "Fleeing Felon, Violating Parole/Probation, Explosives"
				If stat_disq_two_type_code(each_memb) = "26" Then stat_disq_two_type_info(each_memb) = "Family Cash Falsify Residence for Duplicate Assistance"
				If stat_disq_two_type_code(each_memb) = "28" Then stat_disq_two_type_info(each_memb) = "Convicted of Drug Felony - Failed Drug Test"
				If stat_disq_two_type_code(each_memb) = "29" Then stat_disq_two_type_info(each_memb) = "US Citizenship/ID Verif Non-Coop"
				If stat_disq_two_type_code(each_memb) = "30" Then stat_disq_two_type_info(each_memb) = "Immigration Status Verif Non-Coop"
				If stat_disq_two_type_code(each_memb) = "31" Then stat_disq_two_type_info(each_memb) = "EBT Misuse - 1st Disq"
				If stat_disq_two_type_code(each_memb) = "32" Then stat_disq_two_type_info(each_memb) = "EBT Misuse - 2nd Disq"
				If stat_disq_two_type_code(each_memb) = "33" Then stat_disq_two_type_info(each_memb) = "EBT Misuse - 3rd Disq"
				EMReadScreen stat_disq_two_begin_date(each_memb), 8, 8, 64
				stat_disq_two_begin_date(each_memb) = replace(stat_disq_two_begin_date(each_memb), " ", "/")
				If stat_disq_two_begin_date(each_memb) = "__/__/__" Then stat_disq_two_begin_date(each_memb) = ""
				EMReadScreen stat_disq_two_end_date(each_memb), 8, 9, 64
				stat_disq_two_end_date(each_memb) = replace(stat_disq_two_end_date(each_memb), " ", "/")
				If stat_disq_two_end_date(each_memb) = "__/__/__" Then stat_disq_two_end_date(each_memb) = ""
				EMReadScreen stat_disq_two_cure_reason_code(each_memb), 1, 11, 64
				If stat_disq_two_cure_reason_code(each_memb) = "_" Then stat_disq_two_cure_reason_info(each_memb) = ""
				If stat_disq_two_cure_reason_code(each_memb) = "A" Then stat_disq_two_cure_reason_info(each_memb) = "No longer Fleeing Felon/Parole Violation"
				If stat_disq_two_cure_reason_code(each_memb) = "0" Then stat_disq_two_cure_reason_info(each_memb) = "Property Returned/Adequate Compensation"
				If stat_disq_two_cure_reason_code(each_memb) = "1" Then stat_disq_two_cure_reason_info(each_memb) = "Return to Same Job"
				If stat_disq_two_cure_reason_code(each_memb) = "2" Then stat_disq_two_cure_reason_info(each_memb) = "Accept Comparable Employment"
				If stat_disq_two_cure_reason_code(each_memb) = "6" Then stat_disq_two_cure_reason_info(each_memb) = "Cooperate with State QC"
				If stat_disq_two_cure_reason_code(each_memb) = "7" Then stat_disq_two_cure_reason_info(each_memb) = "Cooperate with Federal QC"
				If stat_disq_two_cure_reason_code(each_memb) = "8" Then stat_disq_two_cure_reason_info(each_memb) = "Lump Sum Recalculated"
				EMReadScreen stat_disq_two_fraud_determination_date(each_memb), 8, 13, 64
				stat_disq_two_fraud_determination_date(each_memb) = replace(stat_disq_two_fraud_determination_date(each_memb), " ", "/")
				If stat_disq_two_fraud_determination_date(each_memb) = "__/__/__" Then stat_disq_two_fraud_determination_date(each_memb) = ""
				EMReadScreen stat_disq_two_county_of_fraud(each_memb), 2, 15, 64
				EMReadScreen stat_disq_two_state_of_fraud(each_memb), 2, 16, 64
				EMReadScreen stat_disq_two_SNAP_trafficking_yn(each_memb), 1, 17, 64
				EMReadScreen stat_disq_two_SNAP_offense_code(each_memb), 2, 18, 64
				If stat_disq_two_SNAP_offense_code(each_memb) = "__" Then stat_disq_two_SNAP_offense_info(each_memb) = ""
				If stat_disq_two_SNAP_offense_code(each_memb) = "AL" Then stat_disq_two_SNAP_offense_info(each_memb) = "Alcohol"
				If stat_disq_two_SNAP_offense_code(each_memb) = "DR" Then stat_disq_two_SNAP_offense_info(each_memb) = "Drugs"
				If stat_disq_two_SNAP_offense_code(each_memb) = "GU" Then stat_disq_two_SNAP_offense_info(each_memb) = "Guns"
				If stat_disq_two_SNAP_offense_code(each_memb) = "OT" Then stat_disq_two_SNAP_offense_info(each_memb) = "Other"

				If stat_disq_two_type_code(each_memb) = "02" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "03" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "04" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "15" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "16" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "17" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "18" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "23" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "26" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "28" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "31" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "32" Then stat_disq_two_source(each_memb) = "DISQUAL"
				If stat_disq_two_type_code(each_memb) = "32" Then stat_disq_two_source(each_memb) = "DISQUAL"

				If stat_disq_two_type_code(each_memb) = "06" Then stat_disq_two_source(each_memb) = "NON-COOP"
				If stat_disq_two_type_code(each_memb) = "07" Then stat_disq_two_source(each_memb) = "NON-COOP"
				If stat_disq_two_type_code(each_memb) = "08" Then stat_disq_two_source(each_memb) = "NON-COOP"
				If stat_disq_two_type_code(each_memb) = "14" Then stat_disq_two_source(each_memb) = "NON-COOP"
				If stat_disq_two_type_code(each_memb) = "29" Then stat_disq_two_source(each_memb) = "NON-COOP"
				If stat_disq_two_type_code(each_memb) = "30" Then stat_disq_two_source(each_memb) = "NON-COOP"

				If stat_disq_two_type_code(each_memb) = "12" Then stat_disq_two_source(each_memb) = "TRANSFER"
				If stat_disq_two_type_code(each_memb) = "20" Then stat_disq_two_source(each_memb) = "TRANSFER"

				If stat_disq_two_type_code(each_memb) = "11" Then stat_disq_two_source(each_memb) = "VOL QUIT"

				stat_disq_two_active(each_memb) = True
				If IsDate(stat_disq_two_end_date(each_memb)) = True Then
					If DateDiff("m", stat_disq_two_end_date(each_memb), current_month) >= 0 Then stat_disq_two_active(each_memb) = False
				End If
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "03", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_disq_three_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_disq_three_exists(each_memb) = False

			If stat_disq_three_exists(each_memb) = True Then
				EMReadScreen stat_disq_three_program(each_memb), 2, 6, 54
				If stat_disq_three_program(each_memb) = "__" Then stat_disq_three_program(each_memb) = ""
				If stat_disq_three_program(each_memb) = "AF" Then stat_disq_three_program(each_memb) = "AFDC"
				If stat_disq_three_program(each_memb) = "CC" Then stat_disq_three_program(each_memb) = "Child Care Assistance"
				If stat_disq_three_program(each_memb) = "DW" Then stat_disq_three_program(each_memb) = "DWP"
				If stat_disq_three_program(each_memb) = "FG" Then stat_disq_three_program(each_memb) = "Family General Assistance"
				If stat_disq_three_program(each_memb) = "FS" Then stat_disq_three_program(each_memb) = "SNAP"
				If stat_disq_three_program(each_memb) = "GA" Then stat_disq_three_program(each_memb) = "General Asssistance"
				If stat_disq_three_program(each_memb) = "GR" Then stat_disq_three_program(each_memb) = "GRH"
				If stat_disq_three_program(each_memb) = "IM" Then stat_disq_three_program(each_memb) = "IMD"
				If stat_disq_three_program(each_memb) = "MA" Then stat_disq_three_program(each_memb) = "Medical Assistance"
				If stat_disq_three_program(each_memb) = "MF" Then stat_disq_three_program(each_memb) = "MFIP"
				If stat_disq_three_program(each_memb) = "MS" Then stat_disq_three_program(each_memb) = "MN Supplemental Aid"
				If stat_disq_three_program(each_memb) = "QI" Then stat_disq_three_program(each_memb) = "QI-1"
				If stat_disq_three_program(each_memb) = "QM" Then stat_disq_three_program(each_memb) = "QMB"
				If stat_disq_three_program(each_memb) = "QW" Then stat_disq_three_program(each_memb) = "QWD"
				If stat_disq_three_program(each_memb) = "RC" Then stat_disq_three_program(each_memb) = "Refugee Cash Assistance"
				If stat_disq_three_program(each_memb) = "RM" Then stat_disq_three_program(each_memb) = "RMA"
				If stat_disq_three_program(each_memb) = "SL" Then stat_disq_three_program(each_memb) = "SLMB"
				If stat_disq_three_program(each_memb) = "WB" Then stat_disq_three_program(each_memb) = "Work Benefit Program"
				If stat_disq_three_program(each_memb) = "4E" Then stat_disq_three_program(each_memb) = "Title IV-E Foster Care"
				EMReadScreen stat_disq_three_type_code(each_memb), 2, 6, 64
				If stat_disq_three_type_code(each_memb) = "__" Then stat_disq_three_type_info(each_memb) = ""
				If stat_disq_three_type_code(each_memb) = "02" Then stat_disq_three_type_info(each_memb) = "SNAP Fraud - 1st Disq"
				If stat_disq_three_type_code(each_memb) = "03" Then stat_disq_three_type_info(each_memb) = "SNAP Fraud - 2md Disq"
				If stat_disq_three_type_code(each_memb) = "04" Then stat_disq_three_type_info(each_memb) = "SNAP Fraud - 3rd Disq"
				If stat_disq_three_type_code(each_memb) = "06" Then stat_disq_three_type_info(each_memb) = "Non-Coop with State QC"
				If stat_disq_three_type_code(each_memb) = "07" Then stat_disq_three_type_info(each_memb) = "Non-Coop with Federal QC"
				If stat_disq_three_type_code(each_memb) = "08" Then stat_disq_three_type_info(each_memb) = "RCA Non-Comply with E&T"
				If stat_disq_three_type_code(each_memb) = "11" Then stat_disq_three_type_info(each_memb) = "Voluntary Quit"
				If stat_disq_three_type_code(each_memb) = "12" Then stat_disq_three_type_info(each_memb) = "Improper Transfer of Assets"
				If stat_disq_three_type_code(each_memb) = "13" Then stat_disq_three_type_info(each_memb) = "Lump Sum"
				If stat_disq_three_type_code(each_memb) = "14" Then stat_disq_three_type_info(each_memb) = "IEVS Non-Coop"
				If stat_disq_three_type_code(each_memb) = "15" Then stat_disq_three_type_info(each_memb) = "Cash Fraud - Time Set by Court"
				If stat_disq_three_type_code(each_memb) = "16" Then stat_disq_three_type_info(each_memb) = "Cash Fraud - 1st Disq"
				If stat_disq_three_type_code(each_memb) = "17" Then stat_disq_three_type_info(each_memb) = "Cash Fraud - 2nd Disq"
				If stat_disq_three_type_code(each_memb) = "18" Then stat_disq_three_type_info(each_memb) = "Cash Fraud - 3rd Disq"
				If stat_disq_three_type_code(each_memb) = "20" Then stat_disq_three_type_info(each_memb) = "Improper Transfer of Income"
				If stat_disq_three_type_code(each_memb) = "23" Then stat_disq_three_type_info(each_memb) = "Fleeing Felon, Violating Parole/Probation, Explosives"
				If stat_disq_three_type_code(each_memb) = "26" Then stat_disq_three_type_info(each_memb) = "Family Cash Falsify Residence for Duplicate Assistance"
				If stat_disq_three_type_code(each_memb) = "28" Then stat_disq_three_type_info(each_memb) = "Convicted of Drug Felony - Failed Drug Test"
				If stat_disq_three_type_code(each_memb) = "29" Then stat_disq_three_type_info(each_memb) = "US Citizenship/ID Verif Non-Coop"
				If stat_disq_three_type_code(each_memb) = "30" Then stat_disq_three_type_info(each_memb) = "Immigration Status Verif Non-Coop"
				If stat_disq_three_type_code(each_memb) = "31" Then stat_disq_three_type_info(each_memb) = "EBT Misuse - 1st Disq"
				If stat_disq_three_type_code(each_memb) = "32" Then stat_disq_three_type_info(each_memb) = "EBT Misuse - 2nd Disq"
				If stat_disq_three_type_code(each_memb) = "33" Then stat_disq_three_type_info(each_memb) = "EBT Misuse - 3rd Disq"
				EMReadScreen stat_disq_three_begin_date(each_memb), 8, 8, 64
				stat_disq_three_begin_date(each_memb) = replace(stat_disq_three_begin_date(each_memb), " ", "/")
				If stat_disq_three_begin_date(each_memb) = "__/__/__" Then stat_disq_three_begin_date(each_memb) = ""
				EMReadScreen stat_disq_three_end_date(each_memb), 8, 9, 64
				stat_disq_three_end_date(each_memb) = replace(stat_disq_three_end_date(each_memb), " ", "/")
				If stat_disq_three_end_date(each_memb) = "__/__/__" Then stat_disq_three_end_date(each_memb) = ""
				EMReadScreen stat_disq_three_cure_reason_code(each_memb), 1, 11, 64
				If stat_disq_three_cure_reason_code(each_memb) = "_" Then stat_disq_three_cure_reason_info(each_memb) = ""
				If stat_disq_three_cure_reason_code(each_memb) = "A" Then stat_disq_three_cure_reason_info(each_memb) = "No longer Fleeing Felon/Parole Violation"
				If stat_disq_three_cure_reason_code(each_memb) = "0" Then stat_disq_three_cure_reason_info(each_memb) = "Property Returned/Adequate Compensation"
				If stat_disq_three_cure_reason_code(each_memb) = "1" Then stat_disq_three_cure_reason_info(each_memb) = "Return to Same Job"
				If stat_disq_three_cure_reason_code(each_memb) = "2" Then stat_disq_three_cure_reason_info(each_memb) = "Accept Comparable Employment"
				If stat_disq_three_cure_reason_code(each_memb) = "6" Then stat_disq_three_cure_reason_info(each_memb) = "Cooperate with State QC"
				If stat_disq_three_cure_reason_code(each_memb) = "7" Then stat_disq_three_cure_reason_info(each_memb) = "Cooperate with Federal QC"
				If stat_disq_three_cure_reason_code(each_memb) = "8" Then stat_disq_three_cure_reason_info(each_memb) = "Lump Sum Recalculated"
				EMReadScreen stat_disq_three_fraud_determination_date(each_memb), 8, 13, 64
				stat_disq_three_fraud_determination_date(each_memb) = replace(stat_disq_three_fraud_determination_date(each_memb), " ", "/")
				If stat_disq_three_fraud_determination_date(each_memb) = "__/__/__" Then stat_disq_three_fraud_determination_date(each_memb) = ""
				EMReadScreen stat_disq_three_county_of_fraud(each_memb), 2, 15, 64
				EMReadScreen stat_disq_three_state_of_fraud(each_memb), 2, 16, 64
				EMReadScreen stat_disq_three_SNAP_trafficking_yn(each_memb), 1, 17, 64
				EMReadScreen stat_disq_three_SNAP_offense_code(each_memb), 2, 18, 64
				If stat_disq_three_SNAP_offense_code(each_memb) = "__" Then stat_disq_three_SNAP_offense_info(each_memb) = ""
				If stat_disq_three_SNAP_offense_code(each_memb) = "AL" Then stat_disq_three_SNAP_offense_info(each_memb) = "Alcohol"
				If stat_disq_three_SNAP_offense_code(each_memb) = "DR" Then stat_disq_three_SNAP_offense_info(each_memb) = "Drugs"
				If stat_disq_three_SNAP_offense_code(each_memb) = "GU" Then stat_disq_three_SNAP_offense_info(each_memb) = "Guns"
				If stat_disq_three_SNAP_offense_code(each_memb) = "OT" Then stat_disq_three_SNAP_offense_info(each_memb) = "Other"

				If stat_disq_three_type_code(each_memb) = "02" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "03" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "04" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "15" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "16" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "17" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "18" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "23" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "26" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "28" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "31" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "32" Then stat_disq_three_source(each_memb) = "DISQUAL"
				If stat_disq_three_type_code(each_memb) = "32" Then stat_disq_three_source(each_memb) = "DISQUAL"

				If stat_disq_three_type_code(each_memb) = "06" Then stat_disq_three_source(each_memb) = "NON-COOP"
				If stat_disq_three_type_code(each_memb) = "07" Then stat_disq_three_source(each_memb) = "NON-COOP"
				If stat_disq_three_type_code(each_memb) = "08" Then stat_disq_three_source(each_memb) = "NON-COOP"
				If stat_disq_three_type_code(each_memb) = "14" Then stat_disq_three_source(each_memb) = "NON-COOP"
				If stat_disq_three_type_code(each_memb) = "29" Then stat_disq_three_source(each_memb) = "NON-COOP"
				If stat_disq_three_type_code(each_memb) = "30" Then stat_disq_three_source(each_memb) = "NON-COOP"

				If stat_disq_three_type_code(each_memb) = "12" Then stat_disq_three_source(each_memb) = "TRANSFER"
				If stat_disq_three_type_code(each_memb) = "20" Then stat_disq_three_source(each_memb) = "TRANSFER"

				If stat_disq_three_type_code(each_memb) = "11" Then stat_disq_three_source(each_memb) = "VOL QUIT"

				stat_disq_three_active(each_memb) = True
				If IsDate(stat_disq_three_end_date(each_memb)) = True Then
					If DateDiff("m", stat_disq_three_end_date(each_memb), current_month) >= 0 Then stat_disq_three_active(each_memb) = False
				End If
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "04", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_disq_four_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_disq_four_exists(each_memb) = False

			If stat_disq_four_exists(each_memb) = True Then
				EMReadScreen stat_disq_four_program(each_memb), 2, 6, 54
				If stat_disq_four_program(each_memb) = "__" Then stat_disq_four_program(each_memb) = ""
				If stat_disq_four_program(each_memb) = "AF" Then stat_disq_four_program(each_memb) = "AFDC"
				If stat_disq_four_program(each_memb) = "CC" Then stat_disq_four_program(each_memb) = "Child Care Assistance"
				If stat_disq_four_program(each_memb) = "DW" Then stat_disq_four_program(each_memb) = "DWP"
				If stat_disq_four_program(each_memb) = "FG" Then stat_disq_four_program(each_memb) = "Family General Assistance"
				If stat_disq_four_program(each_memb) = "FS" Then stat_disq_four_program(each_memb) = "SNAP"
				If stat_disq_four_program(each_memb) = "GA" Then stat_disq_four_program(each_memb) = "General Asssistance"
				If stat_disq_four_program(each_memb) = "GR" Then stat_disq_four_program(each_memb) = "GRH"
				If stat_disq_four_program(each_memb) = "IM" Then stat_disq_four_program(each_memb) = "IMD"
				If stat_disq_four_program(each_memb) = "MA" Then stat_disq_four_program(each_memb) = "Medical Assistance"
				If stat_disq_four_program(each_memb) = "MF" Then stat_disq_four_program(each_memb) = "MFIP"
				If stat_disq_four_program(each_memb) = "MS" Then stat_disq_four_program(each_memb) = "MN Supplemental Aid"
				If stat_disq_four_program(each_memb) = "QI" Then stat_disq_four_program(each_memb) = "QI-1"
				If stat_disq_four_program(each_memb) = "QM" Then stat_disq_four_program(each_memb) = "QMB"
				If stat_disq_four_program(each_memb) = "QW" Then stat_disq_four_program(each_memb) = "QWD"
				If stat_disq_four_program(each_memb) = "RC" Then stat_disq_four_program(each_memb) = "Refugee Cash Assistance"
				If stat_disq_four_program(each_memb) = "RM" Then stat_disq_four_program(each_memb) = "RMA"
				If stat_disq_four_program(each_memb) = "SL" Then stat_disq_four_program(each_memb) = "SLMB"
				If stat_disq_four_program(each_memb) = "WB" Then stat_disq_four_program(each_memb) = "Work Benefit Program"
				If stat_disq_four_program(each_memb) = "4E" Then stat_disq_four_program(each_memb) = "Title IV-E Foster Care"
				EMReadScreen stat_disq_four_type_code(each_memb), 2, 6, 64
				If stat_disq_four_type_code(each_memb) = "__" Then stat_disq_four_type_info(each_memb) = ""
				If stat_disq_four_type_code(each_memb) = "02" Then stat_disq_four_type_info(each_memb) = "SNAP Fraud - 1st Disq"
				If stat_disq_four_type_code(each_memb) = "03" Then stat_disq_four_type_info(each_memb) = "SNAP Fraud - 2md Disq"
				If stat_disq_four_type_code(each_memb) = "04" Then stat_disq_four_type_info(each_memb) = "SNAP Fraud - 3rd Disq"
				If stat_disq_four_type_code(each_memb) = "06" Then stat_disq_four_type_info(each_memb) = "Non-Coop with State QC"
				If stat_disq_four_type_code(each_memb) = "07" Then stat_disq_four_type_info(each_memb) = "Non-Coop with Federal QC"
				If stat_disq_four_type_code(each_memb) = "08" Then stat_disq_four_type_info(each_memb) = "RCA Non-Comply with E&T"
				If stat_disq_four_type_code(each_memb) = "11" Then stat_disq_four_type_info(each_memb) = "Voluntary Quit"
				If stat_disq_four_type_code(each_memb) = "12" Then stat_disq_four_type_info(each_memb) = "Improper Transfer of Assets"
				If stat_disq_four_type_code(each_memb) = "13" Then stat_disq_four_type_info(each_memb) = "Lump Sum"
				If stat_disq_four_type_code(each_memb) = "14" Then stat_disq_four_type_info(each_memb) = "IEVS Non-Coop"
				If stat_disq_four_type_code(each_memb) = "15" Then stat_disq_four_type_info(each_memb) = "Cash Fraud - Time Set by Court"
				If stat_disq_four_type_code(each_memb) = "16" Then stat_disq_four_type_info(each_memb) = "Cash Fraud - 1st Disq"
				If stat_disq_four_type_code(each_memb) = "17" Then stat_disq_four_type_info(each_memb) = "Cash Fraud - 2nd Disq"
				If stat_disq_four_type_code(each_memb) = "18" Then stat_disq_four_type_info(each_memb) = "Cash Fraud - 3rd Disq"
				If stat_disq_four_type_code(each_memb) = "20" Then stat_disq_four_type_info(each_memb) = "Improper Transfer of Income"
				If stat_disq_four_type_code(each_memb) = "23" Then stat_disq_four_type_info(each_memb) = "Fleeing Felon, Violating Parole/Probation, Explosives"
				If stat_disq_four_type_code(each_memb) = "26" Then stat_disq_four_type_info(each_memb) = "Family Cash Falsify Residence for Duplicate Assistance"
				If stat_disq_four_type_code(each_memb) = "28" Then stat_disq_four_type_info(each_memb) = "Convicted of Drug Felony - Failed Drug Test"
				If stat_disq_four_type_code(each_memb) = "29" Then stat_disq_four_type_info(each_memb) = "US Citizenship/ID Verif Non-Coop"
				If stat_disq_four_type_code(each_memb) = "30" Then stat_disq_four_type_info(each_memb) = "Immigration Status Verif Non-Coop"
				If stat_disq_four_type_code(each_memb) = "31" Then stat_disq_four_type_info(each_memb) = "EBT Misuse - 1st Disq"
				If stat_disq_four_type_code(each_memb) = "32" Then stat_disq_four_type_info(each_memb) = "EBT Misuse - 2nd Disq"
				If stat_disq_four_type_code(each_memb) = "33" Then stat_disq_four_type_info(each_memb) = "EBT Misuse - 3rd Disq"
				EMReadScreen stat_disq_four_begin_date(each_memb), 8, 8, 64
				stat_disq_four_begin_date(each_memb) = replace(stat_disq_four_begin_date(each_memb), " ", "/")
				If stat_disq_four_begin_date(each_memb) = "__/__/__" Then stat_disq_four_begin_date(each_memb) = ""
				EMReadScreen stat_disq_four_end_date(each_memb), 8, 9, 64
				stat_disq_four_end_date(each_memb) = replace(stat_disq_four_end_date(each_memb), " ", "/")
				If stat_disq_four_end_date(each_memb) = "__/__/__" Then stat_disq_four_end_date(each_memb) = ""
				EMReadScreen stat_disq_four_cure_reason_code(each_memb), 1, 11, 64
				If stat_disq_four_cure_reason_code(each_memb) = "_" Then stat_disq_four_cure_reason_info(each_memb) = ""
				If stat_disq_four_cure_reason_code(each_memb) = "A" Then stat_disq_four_cure_reason_info(each_memb) = "No longer Fleeing Felon/Parole Violation"
				If stat_disq_four_cure_reason_code(each_memb) = "0" Then stat_disq_four_cure_reason_info(each_memb) = "Property Returned/Adequate Compensation"
				If stat_disq_four_cure_reason_code(each_memb) = "1" Then stat_disq_four_cure_reason_info(each_memb) = "Return to Same Job"
				If stat_disq_four_cure_reason_code(each_memb) = "2" Then stat_disq_four_cure_reason_info(each_memb) = "Accept Comparable Employment"
				If stat_disq_four_cure_reason_code(each_memb) = "6" Then stat_disq_four_cure_reason_info(each_memb) = "Cooperate with State QC"
				If stat_disq_four_cure_reason_code(each_memb) = "7" Then stat_disq_four_cure_reason_info(each_memb) = "Cooperate with Federal QC"
				If stat_disq_four_cure_reason_code(each_memb) = "8" Then stat_disq_four_cure_reason_info(each_memb) = "Lump Sum Recalculated"
				EMReadScreen stat_disq_four_fraud_determination_date(each_memb), 8, 13, 64
				stat_disq_four_fraud_determination_date(each_memb) = replace(stat_disq_four_fraud_determination_date(each_memb), " ", "/")
				If stat_disq_four_fraud_determination_date(each_memb) = "__/__/__" Then stat_disq_four_fraud_determination_date(each_memb) = ""
				EMReadScreen stat_disq_four_county_of_fraud(each_memb), 2, 15, 64
				EMReadScreen stat_disq_four_state_of_fraud(each_memb), 2, 16, 64
				EMReadScreen stat_disq_four_SNAP_trafficking_yn(each_memb), 1, 17, 64
				EMReadScreen stat_disq_four_SNAP_offense_code(each_memb), 2, 18, 64
				If stat_disq_four_SNAP_offense_code(each_memb) = "__" Then stat_disq_four_SNAP_offense_info(each_memb) = ""
				If stat_disq_four_SNAP_offense_code(each_memb) = "AL" Then stat_disq_four_SNAP_offense_info(each_memb) = "Alcohol"
				If stat_disq_four_SNAP_offense_code(each_memb) = "DR" Then stat_disq_four_SNAP_offense_info(each_memb) = "Drugs"
				If stat_disq_four_SNAP_offense_code(each_memb) = "GU" Then stat_disq_four_SNAP_offense_info(each_memb) = "Guns"
				If stat_disq_four_SNAP_offense_code(each_memb) = "OT" Then stat_disq_four_SNAP_offense_info(each_memb) = "Other"

				If stat_disq_four_type_code(each_memb) = "02" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "03" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "04" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "15" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "16" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "17" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "18" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "23" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "26" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "28" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "31" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "32" Then stat_disq_four_source(each_memb) = "DISQUAL"
				If stat_disq_four_type_code(each_memb) = "32" Then stat_disq_four_source(each_memb) = "DISQUAL"

				If stat_disq_four_type_code(each_memb) = "06" Then stat_disq_four_source(each_memb) = "NON-COOP"
				If stat_disq_four_type_code(each_memb) = "07" Then stat_disq_four_source(each_memb) = "NON-COOP"
				If stat_disq_four_type_code(each_memb) = "08" Then stat_disq_four_source(each_memb) = "NON-COOP"
				If stat_disq_four_type_code(each_memb) = "14" Then stat_disq_four_source(each_memb) = "NON-COOP"
				If stat_disq_four_type_code(each_memb) = "29" Then stat_disq_four_source(each_memb) = "NON-COOP"
				If stat_disq_four_type_code(each_memb) = "30" Then stat_disq_four_source(each_memb) = "NON-COOP"

				If stat_disq_four_type_code(each_memb) = "12" Then stat_disq_four_source(each_memb) = "TRANSFER"
				If stat_disq_four_type_code(each_memb) = "20" Then stat_disq_four_source(each_memb) = "TRANSFER"

				If stat_disq_four_type_code(each_memb) = "11" Then stat_disq_four_source(each_memb) = "VOL QUIT"

				stat_disq_four_active(each_memb) = True
				If IsDate(stat_disq_four_end_date(each_memb)) = True Then
					If DateDiff("m", stat_disq_four_end_date(each_memb), current_month) >= 0 Then stat_disq_four_active(each_memb) = False
				End If
			End If

			EMWriteScreen stat_memb_ref_numb(each_memb), 20, 76
			EMWriteScreen "05", 20, 79
			transmit
			EMReadScreen existance_check, 14, 24, 13
			stat_disq_five_exists(each_memb) = True
			If existance_check = "DOES NOT EXIST" Then stat_disq_five_exists(each_memb) = False

			If stat_disq_five_exists(each_memb) = True Then
				EMReadScreen stat_disq_five_program(each_memb), 2, 6, 54
				If stat_disq_five_program(each_memb) = "__" Then stat_disq_five_program(each_memb) = ""
				If stat_disq_five_program(each_memb) = "AF" Then stat_disq_five_program(each_memb) = "AFDC"
				If stat_disq_five_program(each_memb) = "CC" Then stat_disq_five_program(each_memb) = "Child Care Assistance"
				If stat_disq_five_program(each_memb) = "DW" Then stat_disq_five_program(each_memb) = "DWP"
				If stat_disq_five_program(each_memb) = "FG" Then stat_disq_five_program(each_memb) = "Family General Assistance"
				If stat_disq_five_program(each_memb) = "FS" Then stat_disq_five_program(each_memb) = "SNAP"
				If stat_disq_five_program(each_memb) = "GA" Then stat_disq_five_program(each_memb) = "General Asssistance"
				If stat_disq_five_program(each_memb) = "GR" Then stat_disq_five_program(each_memb) = "GRH"
				If stat_disq_five_program(each_memb) = "IM" Then stat_disq_five_program(each_memb) = "IMD"
				If stat_disq_five_program(each_memb) = "MA" Then stat_disq_five_program(each_memb) = "Medical Assistance"
				If stat_disq_five_program(each_memb) = "MF" Then stat_disq_five_program(each_memb) = "MFIP"
				If stat_disq_five_program(each_memb) = "MS" Then stat_disq_five_program(each_memb) = "MN Supplemental Aid"
				If stat_disq_five_program(each_memb) = "QI" Then stat_disq_five_program(each_memb) = "QI-1"
				If stat_disq_five_program(each_memb) = "QM" Then stat_disq_five_program(each_memb) = "QMB"
				If stat_disq_five_program(each_memb) = "QW" Then stat_disq_five_program(each_memb) = "QWD"
				If stat_disq_five_program(each_memb) = "RC" Then stat_disq_five_program(each_memb) = "Refugee Cash Assistance"
				If stat_disq_five_program(each_memb) = "RM" Then stat_disq_five_program(each_memb) = "RMA"
				If stat_disq_five_program(each_memb) = "SL" Then stat_disq_five_program(each_memb) = "SLMB"
				If stat_disq_five_program(each_memb) = "WB" Then stat_disq_five_program(each_memb) = "Work Benefit Program"
				If stat_disq_five_program(each_memb) = "4E" Then stat_disq_five_program(each_memb) = "Title IV-E Foster Care"
				EMReadScreen stat_disq_five_type_code(each_memb), 2, 6, 64
				If stat_disq_five_type_code(each_memb) = "__" Then stat_disq_five_type_info(each_memb) = ""
				If stat_disq_five_type_code(each_memb) = "02" Then stat_disq_five_type_info(each_memb) = "SNAP Fraud - 1st Disq"
				If stat_disq_five_type_code(each_memb) = "03" Then stat_disq_five_type_info(each_memb) = "SNAP Fraud - 2md Disq"
				If stat_disq_five_type_code(each_memb) = "04" Then stat_disq_five_type_info(each_memb) = "SNAP Fraud - 3rd Disq"
				If stat_disq_five_type_code(each_memb) = "06" Then stat_disq_five_type_info(each_memb) = "Non-Coop with State QC"
				If stat_disq_five_type_code(each_memb) = "07" Then stat_disq_five_type_info(each_memb) = "Non-Coop with Federal QC"
				If stat_disq_five_type_code(each_memb) = "08" Then stat_disq_five_type_info(each_memb) = "RCA Non-Comply with E&T"
				If stat_disq_five_type_code(each_memb) = "11" Then stat_disq_five_type_info(each_memb) = "Voluntary Quit"
				If stat_disq_five_type_code(each_memb) = "12" Then stat_disq_five_type_info(each_memb) = "Improper Transfer of Assets"
				If stat_disq_five_type_code(each_memb) = "13" Then stat_disq_five_type_info(each_memb) = "Lump Sum"
				If stat_disq_five_type_code(each_memb) = "14" Then stat_disq_five_type_info(each_memb) = "IEVS Non-Coop"
				If stat_disq_five_type_code(each_memb) = "15" Then stat_disq_five_type_info(each_memb) = "Cash Fraud - Time Set by Court"
				If stat_disq_five_type_code(each_memb) = "16" Then stat_disq_five_type_info(each_memb) = "Cash Fraud - 1st Disq"
				If stat_disq_five_type_code(each_memb) = "17" Then stat_disq_five_type_info(each_memb) = "Cash Fraud - 2nd Disq"
				If stat_disq_five_type_code(each_memb) = "18" Then stat_disq_five_type_info(each_memb) = "Cash Fraud - 3rd Disq"
				If stat_disq_five_type_code(each_memb) = "20" Then stat_disq_five_type_info(each_memb) = "Improper Transfer of Income"
				If stat_disq_five_type_code(each_memb) = "23" Then stat_disq_five_type_info(each_memb) = "Fleeing Felon, Violating Parole/Probation, Explosives"
				If stat_disq_five_type_code(each_memb) = "26" Then stat_disq_five_type_info(each_memb) = "Family Cash Falsify Residence for Duplicate Assistance"
				If stat_disq_five_type_code(each_memb) = "28" Then stat_disq_five_type_info(each_memb) = "Convicted of Drug Felony - Failed Drug Test"
				If stat_disq_five_type_code(each_memb) = "29" Then stat_disq_five_type_info(each_memb) = "US Citizenship/ID Verif Non-Coop"
				If stat_disq_five_type_code(each_memb) = "30" Then stat_disq_five_type_info(each_memb) = "Immigration Status Verif Non-Coop"
				If stat_disq_five_type_code(each_memb) = "31" Then stat_disq_five_type_info(each_memb) = "EBT Misuse - 1st Disq"
				If stat_disq_five_type_code(each_memb) = "32" Then stat_disq_five_type_info(each_memb) = "EBT Misuse - 2nd Disq"
				If stat_disq_five_type_code(each_memb) = "33" Then stat_disq_five_type_info(each_memb) = "EBT Misuse - 3rd Disq"
				EMReadScreen stat_disq_five_begin_date(each_memb), 8, 8, 64
				stat_disq_five_begin_date(each_memb) = replace(stat_disq_five_begin_date(each_memb), " ", "/")
				If stat_disq_five_begin_date(each_memb) = "__/__/__" Then stat_disq_five_begin_date(each_memb) = ""
				EMReadScreen stat_disq_five_end_date(each_memb), 8, 9, 64
				stat_disq_five_end_date(each_memb) = replace(stat_disq_five_end_date(each_memb), " ", "/")
				If stat_disq_five_end_date(each_memb) = "__/__/__" Then stat_disq_five_end_date(each_memb) = ""
				EMReadScreen stat_disq_five_cure_reason_code(each_memb), 1, 11, 64
				If stat_disq_five_cure_reason_code(each_memb) = "_" Then stat_disq_five_cure_reason_info(each_memb) = ""
				If stat_disq_five_cure_reason_code(each_memb) = "A" Then stat_disq_five_cure_reason_info(each_memb) = "No longer Fleeing Felon/Parole Violation"
				If stat_disq_five_cure_reason_code(each_memb) = "0" Then stat_disq_five_cure_reason_info(each_memb) = "Property Returned/Adequate Compensation"
				If stat_disq_five_cure_reason_code(each_memb) = "1" Then stat_disq_five_cure_reason_info(each_memb) = "Return to Same Job"
				If stat_disq_five_cure_reason_code(each_memb) = "2" Then stat_disq_five_cure_reason_info(each_memb) = "Accept Comparable Employment"
				If stat_disq_five_cure_reason_code(each_memb) = "6" Then stat_disq_five_cure_reason_info(each_memb) = "Cooperate with State QC"
				If stat_disq_five_cure_reason_code(each_memb) = "7" Then stat_disq_five_cure_reason_info(each_memb) = "Cooperate with Federal QC"
				If stat_disq_five_cure_reason_code(each_memb) = "8" Then stat_disq_five_cure_reason_info(each_memb) = "Lump Sum Recalculated"
				EMReadScreen stat_disq_five_fraud_determination_date(each_memb), 8, 13, 64
				stat_disq_five_fraud_determination_date(each_memb) = replace(stat_disq_five_fraud_determination_date(each_memb), " ", "/")
				If stat_disq_five_fraud_determination_date(each_memb) = "__/__/__" Then stat_disq_five_fraud_determination_date(each_memb) = ""
				EMReadScreen stat_disq_five_county_of_fraud(each_memb), 2, 15, 64
				EMReadScreen stat_disq_five_state_of_fraud(each_memb), 2, 16, 64
				EMReadScreen stat_disq_five_SNAP_trafficking_yn(each_memb), 1, 17, 64
				EMReadScreen stat_disq_five_SNAP_offense_code(each_memb), 2, 18, 64
				If stat_disq_five_SNAP_offense_code(each_memb) = "__" Then stat_disq_five_SNAP_offense_info(each_memb) = ""
				If stat_disq_five_SNAP_offense_code(each_memb) = "AL" Then stat_disq_five_SNAP_offense_info(each_memb) = "Alcohol"
				If stat_disq_five_SNAP_offense_code(each_memb) = "DR" Then stat_disq_five_SNAP_offense_info(each_memb) = "Drugs"
				If stat_disq_five_SNAP_offense_code(each_memb) = "GU" Then stat_disq_five_SNAP_offense_info(each_memb) = "Guns"
				If stat_disq_five_SNAP_offense_code(each_memb) = "OT" Then stat_disq_five_SNAP_offense_info(each_memb) = "Other"

				If stat_disq_five_type_code(each_memb) = "02" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "03" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "04" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "15" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "16" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "17" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "18" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "23" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "26" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "28" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "31" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "32" Then stat_disq_five_source(each_memb) = "DISQUAL"
				If stat_disq_five_type_code(each_memb) = "32" Then stat_disq_five_source(each_memb) = "DISQUAL"

				If stat_disq_five_type_code(each_memb) = "06" Then stat_disq_five_source(each_memb) = "NON-COOP"
				If stat_disq_five_type_code(each_memb) = "07" Then stat_disq_five_source(each_memb) = "NON-COOP"
				If stat_disq_five_type_code(each_memb) = "08" Then stat_disq_five_source(each_memb) = "NON-COOP"
				If stat_disq_five_type_code(each_memb) = "14" Then stat_disq_five_source(each_memb) = "NON-COOP"
				If stat_disq_five_type_code(each_memb) = "29" Then stat_disq_five_source(each_memb) = "NON-COOP"
				If stat_disq_five_type_code(each_memb) = "30" Then stat_disq_five_source(each_memb) = "NON-COOP"

				If stat_disq_five_type_code(each_memb) = "12" Then stat_disq_five_source(each_memb) = "TRANSFER"
				If stat_disq_five_type_code(each_memb) = "20" Then stat_disq_five_source(each_memb) = "TRANSFER"

				If stat_disq_five_type_code(each_memb) = "11" Then stat_disq_five_source(each_memb) = "VOL QUIT"

				stat_disq_five_active(each_memb) = True
				If IsDate(stat_disq_five_end_date(each_memb)) = True Then
					If DateDiff("m", stat_disq_five_end_date(each_memb), current_month) >= 0 Then stat_disq_five_active(each_memb) = False
				End If
			End If
		Next

		Call navigate_to_MAXIS_screen("STAT", "HEST")

		EMReadScreen hest_version, 1, 2, 73
		If hest_version = "1" Then
			EMReadScreen stat_hest_persons_paying_list, 29, 6, 40
			stat_hest_persons_paying_list = replace(stat_hest_persons_paying_list, " __", "")
			stat_hest_persons_paying_list = replace(stat_hest_persons_paying_list, " ", ", ")

			EMReadScreen stat_hest_retro_heat_air_yn, 		1, 13, 34
			EMReadScreen stat_hest_retro_heat_air_fs_units, 2, 13, 42
			EMReadScreen stat_hest_retro_heat_air_amount,	6, 13, 49
			EMReadScreen stat_hest_retro_electric_yn, 		1, 14, 34
			EMReadScreen stat_hest_retro_electric_fs_units, 2, 14, 42
			EMReadScreen stat_hest_retro_electric_amount,	6, 14, 49
			EMReadScreen stat_hest_retro_phone_yn, 			1, 15, 34
			EMReadScreen stat_hest_retro_phone_fs_units, 	2, 15, 42
			EMReadScreen stat_hest_retro_phone_amount,		6, 15, 49

			EMReadScreen stat_hest_prosp_heat_air_yn, 		1, 13, 60
			EMReadScreen stat_hest_prosp_heat_air_fs_units, 2, 13, 68
			EMReadScreen stat_hest_prosp_heat_air_amount,	6, 13, 75
			EMReadScreen stat_hest_prosp_electric_yn, 		1, 14, 60
			EMReadScreen stat_hest_prosp_electric_fs_units, 2, 14, 68
			EMReadScreen stat_hest_prosp_electric_amount,	6, 14, 75
			EMReadScreen stat_hest_prosp_phone_yn, 			1, 15, 60
			EMReadScreen stat_hest_prosp_phone_fs_units, 	2, 15, 68
			EMReadScreen stat_hest_prosp_phone_amount,		6, 15, 75

			stat_hest_retro_heat_air_amount = trim(stat_hest_retro_heat_air_amount)
			If stat_hest_retro_heat_air_amount = "" Then stat_hest_retro_heat_air_amount = 0
			stat_hest_retro_heat_air_amount = stat_hest_retro_heat_air_amount*1

			stat_hest_retro_electric_amount = trim(stat_hest_retro_electric_amount)
			If stat_hest_retro_electric_amount = "" Then stat_hest_retro_electric_amount = 0
			stat_hest_retro_electric_amount = stat_hest_retro_electric_amount*1

			stat_hest_retro_phone_amount = trim(stat_hest_retro_phone_amount)
			If stat_hest_retro_phone_amount = "" Then stat_hest_retro_phone_amount = 0
			stat_hest_retro_phone_amount = stat_hest_retro_phone_amount*1

			stat_hest_prosp_heat_air_amount = trim(stat_hest_prosp_heat_air_amount)
			If stat_hest_prosp_heat_air_amount = "" Then stat_hest_prosp_heat_air_amount = 0
			stat_hest_prosp_heat_air_amount = stat_hest_prosp_heat_air_amount*1

			stat_hest_prosp_electric_amount = trim(stat_hest_prosp_electric_amount)
			If stat_hest_prosp_electric_amount = "" Then stat_hest_prosp_electric_amount = 0
			stat_hest_prosp_electric_amount = stat_hest_prosp_electric_amount*1

			stat_hest_prosp_phone_amount = trim(stat_hest_prosp_phone_amount)
			If stat_hest_prosp_phone_amount = "" Then stat_hest_prosp_phone_amount = 0
			stat_hest_prosp_phone_amount = stat_hest_prosp_phone_amount*1

			stat_hest_retro_all = stat_hest_retro_heat_air_amount + stat_hest_retro_electric_amount + stat_hest_retro_phone_amount
			stat_hest_prosp_all = stat_hest_prosp_heat_air_amount + stat_hest_prosp_electric_amount + stat_hest_prosp_phone_amount
			If stat_hest_retro_heat_air_yn = "Y" Then
				stat_hest_retro_list = "Heat/Air Conditioning"
			ElseIf stat_hest_retro_electric_yn = "Y" AND stat_hest_retro_phone_yn = "Y" Then
				stat_hest_retro_list = "Electric and Phone"
			ElseIf stat_hest_retro_electric_yn = "Y" Then
				stat_hest_retro_list = "Electric"
			ElseIf stat_hest_retro_phone_yn = "Y" Then
				stat_hest_retro_list = "Phone"
			Else
				stat_hest_retro_list = "None"
			End If

			If stat_hest_retro_heat_air_yn = "Y" Then
				stat_hest_prosp_list = "Heat/Air Conditioning"
			ElseIf stat_hest_retro_electric_yn = "Y" AND stat_hest_retro_phone_yn = "Y" Then
				stat_hest_prosp_list = "Electric and Phone"
			ElseIf stat_hest_retro_electric_yn = "Y" Then
				stat_hest_prosp_list = "Electric"
			ElseIf stat_hest_retro_phone_yn = "Y" Then
				stat_hest_prosp_list = "Phone"
			Else
				stat_hest_prosp_list = "None"
			End If

		End if

		Call navigate_to_MAXIS_screen("STAT", "REVW")

		EMReadScreen stat_revw_cash_code, 1, 7, 40
		EMReadScreen stat_next_cash_revw_date, 8, 9, 37
		EMReadScreen stat_next_cash_revw_process, 2, 9, 46
		EMReadScreen stat_last_cash_revw_date, 8, 11, 37
		EMReadScreen stat_last_cash_revw_process, 2, 11, 46
		EMReadScreen stat_revw_snap_code, 1, 7, 60
		EMReadScreen stat_next_snap_revw_date, 8, 9, 57
		EMReadScreen stat_next_snap_revw_process, 2, 9, 66
		EMReadScreen stat_last_snap_revw_date, 8, 11, 57
		EMReadScreen stat_last_snap_revw_process, 2, 11, 66
		EMReadScreen stat_revw_hc_code, 1, 7, 73
		EMReadScreen stat_next_hc_revw_date, 8, 9, 70
		EMReadScreen stat_next_hc_revw_process, 2, 9, 79
		EMReadScreen stat_last_hc_revw_date, 8, 11, 70
		EMReadScreen stat_last_hc_revw_process, 2, 11, 79
		EMReadScreen stat_revw_form_recvd_date, 8, 13, 37
		EMReadScreen stat_revw_interview_date, 8, 15, 37

		stat_revw_cash_code = replace(stat_revw_cash_code, "_", "")
		stat_revw_snap_code = replace(stat_revw_snap_code, "_", "")
		stat_revw_hc_code = replace(stat_revw_hc_code, "_", "")

		stat_next_cash_revw_process = trim(stat_next_cash_revw_process)
		stat_last_cash_revw_process = trim(stat_last_cash_revw_process)

		stat_next_snap_revw_process = trim(stat_next_snap_revw_process)
		stat_last_snap_revw_process = trim(stat_last_snap_revw_process)

		stat_next_hc_revw_process = trim(stat_next_hc_revw_process)
		stat_last_hc_revw_process = trim(stat_last_hc_revw_process)

		stat_next_cash_revw_date = replace(stat_next_cash_revw_date, " ", "/")
		If stat_next_cash_revw_date = "__/__/__" Then stat_next_cash_revw_date = ""
		stat_last_cash_revw_date = replace(stat_last_cash_revw_date, " ", "/")
		If stat_last_cash_revw_date = "__/__/__" Then stat_last_cash_revw_date = ""

		stat_next_snap_revw_date = replace(stat_next_snap_revw_date, " ", "/")
		If stat_next_snap_revw_date = "__/__/__" Then stat_next_snap_revw_date = ""
		stat_last_snap_revw_date = replace(stat_last_snap_revw_date, " ", "/")
		If stat_last_snap_revw_date = "__/__/__" Then stat_last_snap_revw_date = ""

		stat_next_hc_revw_date = replace(stat_next_hc_revw_date, " ", "/")
		If stat_next_hc_revw_date = "__/__/__" Then stat_next_hc_revw_date = ""
		stat_last_hc_revw_date = replace(stat_last_hc_revw_date, " ", "/")
		If stat_last_hc_revw_date = "__/__/__" Then stat_last_hc_revw_date = ""

		stat_revw_form_recvd_date = replace(stat_revw_form_recvd_date, " ", "/")
		If stat_revw_form_recvd_date = "__/__/__" Then stat_revw_form_recvd_date = ""
		stat_revw_interview_date = replace(stat_revw_interview_date, " ", "/")
		If stat_revw_interview_date = "__/__/__" Then stat_revw_interview_date = ""


		Call navigate_to_MAXIS_screen("STAT", "MONT")

		EMReadScreen stat_mont_cash_status, 1, 11, 43
		EMReadScreen stat_mont_snap_status, 1, 11, 53
		EMReadScreen stat_mont_hc_status, 1, 11, 63
		EMReadScreen stat_mont_form_recvd_date, 8, 6, 39

		stat_mont_cash_status = replace(stat_mont_cash_status, "_", "")
		stat_mont_snap_status = replace(stat_mont_snap_status, "_", "")
		stat_mont_hc_status = replace(stat_mont_hc_status, "_", "")

		stat_mont_form_recvd_date = replace(stat_mont_form_recvd_date, " ", "/")
		If stat_mont_form_recvd_date = "__/__/__" Then stat_mont_form_recvd_date = ""

		Call back_to_SELF
	end sub
end class
curr_month_plus_one = CM_plus_1_mo & "/" & CM_plus_1_yr

'Constants
const ref_numb_const				= 0

const access_denied					= 1
const full_name_const				= 2
const last_name_const				= 3
const first_name_const				= 4
const mid_initial					= 5
const other_names					= 6
const age							= 7
const date_of_birth					= 8
const ssn							= 9
const ssn_verif						= 10
const birthdate_verif				= 11

const snap_elig_indicator			= 12
const mfip_elig_indicator			= 13


' const fs_request_yn_const			= 12
' const fs_memb_code_const			= 13
' const fs_memb_status_info_const		= 14
' const fs_memb_counted_const			= 15
' const fs_memb_state_food_const		= 16
' const fs_memb_elig_status_const		= 17
' const fs_memb_begin_date_const		= 18
' const fs_memb_budg_cycle_const		= 19
' const fs_memb_abawd_const			= 20
' const fs_memb_absence_const			= 21
' const fs_memb_roomer_const			= 22
' const fs_memb_boarder_const			= 23
' const fs_memb_citizenship_const		= 24
' const fs_memb_citizenship_coop_const = 25
' const fs_memb_cmdty_const			= 26
' const fs_memb_disq_const			= 27
' const fs_memb_dupl_assist_const		= 28
' const fs_memb_fraud_const			= 29
' const fs_memb_eligible_student_const = 30
' const fs_memb_institution_const		= 31
' const fs_memb_mfip_elig_const		= 32
' const fs_memb_non_applcnt_const		= 33
' const fs_memb_residence_const		= 34
' const fs_memb_ssn_coop_const		= 35
' const fs_memb_unit_memb_const		= 36
' const fs_memb_work_reg_const		= 37
' const fs_memb_drug_felon_test_const	= 38

const last_const = 50

'Arrays
Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(last_const, 0)

Dim DWP_ELIG_APPROVALS()
ReDim DWP_ELIG_APPROVALS(0)

Dim MFIP_ELIG_APPROVALS()
ReDim MFIP_ELIG_APPROVALS(0)

Dim MSA_ELIG_APPROVALS()
ReDim MSA_ELIG_APPROVALS(0)

Dim GA_ELIG_APPROVALS()
ReDim GA_ELIG_APPROVALS(0)

Dim CASH_DENIAL_APPROVALS()
ReDim CASH_DENIAL_APPROVALS(0)

Dim GRH_ELIG_APPROVALS()
ReDim GRH_ELIG_APPROVALS(0)

Dim IVE_ELIG_APPROVALS()
ReDim IVE_ELIG_APPROVALS(0)

' Dim EMER_ELIG_APPROVALS()
' ReDim EMER_ELIG_APPROVALS(0)

Dim SNAP_ELIG_APPROVALS()
ReDim SNAP_ELIG_APPROVALS(0)

Dim HC_ELIG_APPROVALS()
ReDim HC_ELIG_APPROVALS(0)

Dim STAT_INFORMATION()
ReDim STAT_INFORMATION(0)

'===========================================================================================================================
EMConnect ""
Call check_for_MAXIS(True)
testing_run = True
If user_ID_for_validation = "AMST002" Then testing_run = False
end_msg_info = ""

Call MAXIS_case_number_finder(MAXIS_case_number)

Do
	Do
		err_msg = ""

		BeginDialog Dialog1, 0, 0, 366, 135, "Eligibility Summary Case Number Dialog"
		  EditBox 75, 10, 60, 15, MAXIS_case_number
		  EditBox 100, 30, 15, 15, first_footer_month
		  EditBox 120, 30, 15, 15, first_footer_year
		  EditBox 10, 65, 125, 15, worker_signature
		  ButtonGroup ButtonPressed
		    OkButton 250, 110, 50, 15
		    CancelButton 305, 110, 50, 15
		    PushButton 250, 60, 105, 15, "Script Instructions", intructions_btn
		  Text 25, 15, 50, 10, "Case Number"
		  Text 30, 35, 65, 10, "First month of APP"
		  Text 105, 45, 35, 10, "MM    YY"
		  Text 10, 55, 80, 10, "Sign your case note(s):"
		  Text 10, 90, 160, 10, "This script does not have an open 'Notes' field."
		  Text 10, 105, 235, 20, "If there were otherr actions/updates to the case, a separete NOTE should be entered (or another script run) to detail the specifics of that action."
		  Text 155, 5, 140, 20, "This script will detail information about all APP actions for a this case taken today."
		  Text 160, 25, 185, 10, "- Script will handle for approvals, denials, and closures."
		  Text 160, 35, 155, 10, "- Script will handle for any program in MAXIS."
		  Text 160, 45, 180, 10, "- To be handled by the script ELIG resulsts must be:"
		  Text 175, 55, 60, 10, "CREATED Today"
		  Text 175, 65, 65, 10, "APPROVED Today"
		EndDialog

		dialog Dialog1

		cancel_without_confirmation

		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(first_footer_month, first_footer_year, err_msg, "*")
		If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your name to sign your case note."

		If ButtonPressed = intructions_btn Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20ELIGIBILITY%20SUMMARY.docx"
		Else
			If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg
		End If

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Eligibility Summary Script cancelled as it was run in inquiry.")
End If

Call date_array_generator(first_footer_month, first_footer_year, MONTHS_ARRAY)

first_DWP_approval = ""
first_MFIP_approval = ""
first_MSA_approval = ""
first_GA_approval = ""
first_DENY_approval = ""
first_GRH_approval = ""
first_SNAP_approval = ""
first_HC_approval = ""

enter_CNOTE_for_DWP = False
enter_CNOTE_for_MFIP = False
enter_CNOTE_for_MSA = False
enter_CNOTE_for_GA = False
enter_CNOTE_for_DENY = False
enter_CNOTE_for_GRH = False
enter_CNOTE_for_EMER = False
enter_CNOTE_for_SNAP = False
enter_CNOTE_for_HC = False

dwp_elig_months_count = 0
mfip_elig_months_count = 0
msa_elig_months_count = 0
ga_elig_months_count = 0
cash_deny_months_count = 0
grh_elig_months_count = 0
' ive_elig_months_count = 0
emer_elig_months_count = 0
snap_elig_months_count = 0
hc_elig_months_count = 0
month_count = 0

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
Call Navigate_to_MAXIS_screen("ELIG", "SUMM")
EMReadScreen numb_EMER_versions, 1, 16, 40

'TODO figure out EMER Date handling'
If numb_EMER_versions <> " " Then
	Set EMER_ELIG_APPROVAL = new emer_eligibility_detail
	EMER_ELIG_APPROVAL.elig_footer_month = CM_mo
	EMER_ELIG_APPROVAL.elig_footer_year = CM_yr

	EMER_ELIG_APPROVAL.initial_search_month = first_footer_month
	EMER_ELIG_APPROVAL.initial_search_year = first_footer_year

	EMER_ELIG_APPROVAL.read_elig

	If EMER_ELIG_APPROVAL.approved_today = True then enter_CNOTE_for_EMER = True
	' transactions = ""
	' for each_tx = 0 to UBound(EMER_ELIG_APPROVAL.emer_check_program)
	' 	transactions = transactions & EMER_ELIG_APPROVAL.emer_check_program(each_tx) & " - $" & EMER_ELIG_APPROVAL.emer_check_transaction_amount(each_tx) & " Paid to: " & EMER_ELIG_APPROVAL.emer_check_vendor_name(each_tx)
	' 	transactions = transactions & vbCr
	' Next
	'
	' MsgBox "EMER_ELIG_APPROVAL.elig_footer_month - " & EMER_ELIG_APPROVAL.elig_footer_month & vbCr & "EMER_ELIG_APPROVAL.elig_footer_year - " & EMER_ELIG_APPROVAL.elig_footer_year & vbCr &_
	' 		"EMER_ELIG_APPROVAL.emer_elig_summ_eligibility_result - " & EMER_ELIG_APPROVAL.emer_elig_summ_eligibility_result & vbCr &_
	' 		"EMER_ELIG_APPROVAL.emer_elig_summ_payment - " & EMER_ELIG_APPROVAL.emer_elig_summ_payment & vbCr &_
	' 		transactions
End If

For each footer_month in MONTHS_ARRAY
	' MsgBox footer_month
	Call convert_date_into_MAXIS_footer_month(footer_month, MAXIS_footer_month, MAXIS_footer_year)

	ReDim preserve STAT_INFORMATION(month_count)

	Set STAT_INFORMATION(month_count) = new stat_detail

	STAT_INFORMATION(month_count).footer_month = MAXIS_footer_month
	STAT_INFORMATION(month_count).footer_year = MAXIS_footer_year

	Call STAT_INFORMATION(month_count).gather_stat_info


	Call Navigate_to_MAXIS_screen("ELIG", "SUMM")

	EMReadScreen numb_DWP_versions, 		1, 7, 40
	EMReadScreen numb_MFIP_versions, 		1, 8, 40
	EMReadScreen numb_MSA_versions, 		1, 11, 40
	EMReadScreen numb_GA_versions, 			1, 12, 40
	EMReadScreen numb_CASH_denial_versions, 1, 13, 40
	EMReadScreen numb_GRH_versions, 		1, 14, 40
	' EMReadScreen numb_IVE_versions, 		1, 15, 40
	' EMReadScreen numb_EMER_versions, 		1, 16, 40		- WE WILL NOT LOOK AT THIS EVERY MONTH
	EMReadScreen numb_SNAP_versions, 		1, 17, 40

	' MsgBox "numb_SNAP_versions - " & numb_SNAP_versions
	'TODO MAKE THIS READ THE DATE AND COMPARE TO TODAY

	If numb_DWP_versions <> " " Then
		ReDim Preserve DWP_ELIG_APPROVALS(dwp_elig_months_count)
		Set DWP_ELIG_APPROVALS(dwp_elig_months_count) = new dwp_eligibility_detail

		DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_month = MAXIS_footer_month
		DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call DWP_ELIG_APPROVALS(dwp_elig_months_count).read_elig

		If first_DWP_approval = "" AND DWP_ELIG_APPROVALS(dwp_elig_months_count).approved_today Then first_DWP_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
		' MsgBox "DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_month - " & DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_month & vbCr & "DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_year - " & DWP_ELIG_APPROVALS(dwp_elig_months_count).elig_footer_year & vbCr &_
		' "DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_approved_date: " & DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_approved_date & vbCr & "DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_summary_grant_amount: " & DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_summary_grant_amount & vbCr &_
		' "DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_eligibility_result: " & DWP_ELIG_APPROVALS(dwp_elig_months_count).dwp_case_eligibility_result

		dwp_elig_months_count = dwp_elig_months_count + 1
	End If

	If numb_MFIP_versions <> " " Then
		' MsgBox "In MFIP"
		ReDim Preserve MFIP_ELIG_APPROVALS(mfip_elig_months_count)
		Set MFIP_ELIG_APPROVALS(mfip_elig_months_count) = new mfip_eligibility_detail

		MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_month = MAXIS_footer_month
		MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call MFIP_ELIG_APPROVALS(mfip_elig_months_count).read_elig

		MFIP_ELIG_APPROVALS(mfip_elig_months_count).er_month = False
		MFIP_ELIG_APPROVALS(mfip_elig_months_count).hrf_month = False

		If STAT_INFORMATION(month_count).stat_revw_cash_code <> "" Then
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).er_month = True
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).er_status = STAT_INFORMATION(month_count).stat_revw_cash_code
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).er_caf_date = STAT_INFORMATION(month_count).stat_revw_form_recvd_date
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).er_interview_date = STAT_INFORMATION(month_count).stat_revw_interview_date
		End If
		If STAT_INFORMATION(month_count).stat_mont_cash_status <> "" Then
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).hrf_month = True
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).hrf_status = STAT_INFORMATION(month_count).stat_mont_cash_status
			MFIP_ELIG_APPROVALS(mfip_elig_months_count).hrf_doc_date = STAT_INFORMATION(month_count).stat_mont_form_recvd_date
		End If

		If first_MFIP_approval = "" AND MFIP_ELIG_APPROVALS(mfip_elig_months_count).approved_today Then first_MFIP_approval = MAXIS_footer_month & "/" & MAXIS_footer_year

		' MsgBox "MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_month - " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_month & vbCr & "MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_year - " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).elig_footer_year & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_approved_date: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_approved_date & vbCr & "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_grant_amount: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_grant_amount & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_cash_portion: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_cash_portion & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_food_portion: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_food_portion & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_housing_grant: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_summary_housing_grant & vbCr &_
		' "MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_eligibility_result: " & MFIP_ELIG_APPROVALS(mfip_elig_months_count).mfip_case_eligibility_result

		mfip_elig_months_count = mfip_elig_months_count + 1
		' MsgBox "mfip_elig_months_count: " & mfip_elig_months_count
	End If

	If numb_MSA_versions <> " " Then
		ReDim Preserve MSA_ELIG_APPROVALS(msa_elig_months_count)
		Set MSA_ELIG_APPROVALS(msa_elig_months_count) = new msa_eligibility_detail

		MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_month = MAXIS_footer_month
		MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call MSA_ELIG_APPROVALS(msa_elig_months_count).read_elig

		MSA_ELIG_APPROVALS(msa_elig_months_count).er_month = False

		If STAT_INFORMATION(month_count).stat_revw_cash_code <> "" Then
			MSA_ELIG_APPROVALS(msa_elig_months_count).er_month = True
			MSA_ELIG_APPROVALS(msa_elig_months_count).er_status = STAT_INFORMATION(month_count).stat_revw_cash_code
			MSA_ELIG_APPROVALS(msa_elig_months_count).er_caf_date = STAT_INFORMATION(month_count).stat_revw_form_recvd_date
			MSA_ELIG_APPROVALS(msa_elig_months_count).er_interview_date = STAT_INFORMATION(month_count).stat_revw_interview_date
		End If

		If first_MSA_approval = "" AND MSA_ELIG_APPROVALS(msa_elig_months_count).approved_today Then first_MSA_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
		' MsgBox "MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_month - " & MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_month & vbCr & "MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_year - " & MSA_ELIG_APPROVALS(msa_elig_months_count).elig_footer_year & vbCr &_
		' "MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_approved_date: " & MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_approved_date & vbCr & "MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_grant: " & MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_grant & vbCr &_
		' "MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_eligibility_result: " & MSA_ELIG_APPROVALS(msa_elig_months_count).msa_elig_summ_eligibility_result

		msa_elig_months_count = msa_elig_months_count + 1
	End If

	If numb_GA_versions <> " " Then
		ReDim Preserve GA_ELIG_APPROVALS(ga_elig_months_count)
		Set GA_ELIG_APPROVALS(ga_elig_months_count) = new ga_eligibility_detail

		GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_month = MAXIS_footer_month
		GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call GA_ELIG_APPROVALS(ga_elig_months_count).read_elig

		GA_ELIG_APPROVALS(ga_elig_months_count).er_month = False
		GA_ELIG_APPROVALS(ga_elig_months_count).hrf_month = False

		If STAT_INFORMATION(month_count).stat_revw_cash_code <> "" Then
			GA_ELIG_APPROVALS(ga_elig_months_count).er_month = True
			GA_ELIG_APPROVALS(ga_elig_months_count).er_status = STAT_INFORMATION(month_count).stat_revw_cash_code
			GA_ELIG_APPROVALS(ga_elig_months_count).er_caf_date = STAT_INFORMATION(month_count).stat_revw_form_recvd_date
			GA_ELIG_APPROVALS(ga_elig_months_count).er_interview_date = STAT_INFORMATION(month_count).stat_revw_interview_date
		End If
		If STAT_INFORMATION(month_count).stat_mont_cash_status <> "" Then
			GA_ELIG_APPROVALS(ga_elig_months_count).hrf_month = True
			GA_ELIG_APPROVALS(ga_elig_months_count).hrf_status = STAT_INFORMATION(month_count).stat_mont_cash_status
			GA_ELIG_APPROVALS(ga_elig_months_count).hrf_doc_date = STAT_INFORMATION(month_count).stat_mont_form_recvd_date
		End If

		If first_GA_approval = "" AND GA_ELIG_APPROVALS(ga_elig_months_count).approved_today Then first_GA_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
		' MsgBox "GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_month - " & GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_month & vbCr & "GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_year - " & GA_ELIG_APPROVALS(ga_elig_months_count).elig_footer_year & vbCr &_
		' "GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_approved_date: " & GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_approved_date & vbCr & "GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_monthly_grant: " & GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_monthly_grant & vbCr &_
		' "GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_eligibility_result: " & GA_ELIG_APPROVALS(ga_elig_months_count).ga_elig_summ_eligibility_result

		ga_elig_months_count = ga_elig_months_count + 1
	End If

	If numb_CASH_denial_versions <> " " Then
		ReDim Preserve CASH_DENIAL_APPROVALS(cash_deny_months_count)
		Set CASH_DENIAL_APPROVALS(cash_deny_months_count) = new deny_eligibility_detail

		CASH_DENIAL_APPROVALS(cash_deny_months_count).elig_footer_month = MAXIS_footer_month
		CASH_DENIAL_APPROVALS(cash_deny_months_count).elig_footer_year = MAXIS_footer_year

		Call CASH_DENIAL_APPROVALS(cash_deny_months_count).read_elig

		If first_DENY_approval = "" AND CASH_DENIAL_APPROVALS(cash_deny_months_count).approved_today Then first_DENY_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
		' members = ""
		' for each_memb = 0 to UBound(CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_membs_ref_numbs)
		' 	members = members & "MEMB " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_membs_ref_numbs(each_memb) & " - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_membs_full_name(each_memb) & " Request: " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_membs_request_yn(each_memb)
		' 	members = members & vbCr
		' Next

		' MsgBox "CASH_DENIAL_APPROVALS(cash_deny_months_count).elig_footer_month - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).elig_footer_month & vbCr & "CASH_DENIAL_APPROVALS(cash_deny_months_count).elig_footer_year - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).elig_footer_year & vbCr &_
		' "CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_dwp_reason_info - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_dwp_reason_info & vbCr &_
		' "CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_mfip_reason_info - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_mfip_reason_info & vbCr &_
		' "CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_msa_reason_info - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_msa_reason_info & vbCr &_
		' "CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_ga_reason_info - " & CASH_DENIAL_APPROVALS(cash_deny_months_count).deny_cash_ga_reason_info & vbCr &_
		' members

		cash_deny_months_count = cash_deny_months_count + 1
	End If

	If numb_GRH_versions <> " " Then
		ReDim Preserve GRH_ELIG_APPROVALS(grh_elig_months_count)
		Set GRH_ELIG_APPROVALS(grh_elig_months_count) = new grh_eligibility_detail

		GRH_ELIG_APPROVALS(grh_elig_months_count).elig_footer_month = MAXIS_footer_month
		GRH_ELIG_APPROVALS(grh_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call GRH_ELIG_APPROVALS(grh_elig_months_count).read_elig

		GRH_ELIG_APPROVALS(grh_elig_months_count).er_month = False
		GRH_ELIG_APPROVALS(grh_elig_months_count).hrf_month = False

		If STAT_INFORMATION(month_count).stat_revw_cash_code <> "" Then
			GRH_ELIG_APPROVALS(grh_elig_months_count).er_month = True
			GRH_ELIG_APPROVALS(grh_elig_months_count).er_status = STAT_INFORMATION(month_count).stat_revw_cash_code
			GRH_ELIG_APPROVALS(grh_elig_months_count).er_caf_date = STAT_INFORMATION(month_count).stat_revw_form_recvd_date
			GRH_ELIG_APPROVALS(grh_elig_months_count).er_interview_date = STAT_INFORMATION(month_count).stat_revw_interview_date
		End If
		If STAT_INFORMATION(month_count).stat_mont_cash_status <> "" Then
			GRH_ELIG_APPROVALS(grh_elig_months_count).hrf_month = True
			GRH_ELIG_APPROVALS(grh_elig_months_count).hrf_status = STAT_INFORMATION(month_count).stat_mont_cash_status
			GRH_ELIG_APPROVALS(grh_elig_months_count).hrf_doc_date = STAT_INFORMATION(month_count).stat_mont_form_recvd_date
		End If

		If first_GRH_approval = "" AND GRH_ELIG_APPROVALS(grh_elig_months_count).approved_today Then first_GRH_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
		' MsgBox "GRH_ELIG_APPROVALS(grh_elig_months_count).elig_footer_month - " & GRH_ELIG_APPROVALS(grh_elig_months_count).elig_footer_month & vbCr & "GRH_ELIG_APPROVALS(grh_elig_months_count).elig_footer_year - " & GRH_ELIG_APPROVALS(grh_elig_months_count).elig_footer_year & vbCr &_
		' "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_approved_date: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_approved_date & vbCr &_
		' "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_payable_amount_one: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_payable_amount_one & vbCr &_
		' "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_vendor_one_name: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_vendor_one_name & vbCr & "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_budg_vendor_number_one: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_budg_vendor_number_one & vbCr &_
		' "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_payable_amount_two: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_payable_amount_two & vbCr &_
		' "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_vendor_two_name: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_vendor_two_name & vbCr & "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_budg_vendor_number_two: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_budg_vendor_number_two & vbCr &_
		' "GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_eligibility_result: " & GRH_ELIG_APPROVALS(grh_elig_months_count).grh_elig_eligibility_result

		grh_elig_months_count = grh_elig_months_count + 1
	End If

	If numb_SNAP_versions <> " " Then
		ReDim Preserve SNAP_ELIG_APPROVALS(snap_elig_months_count)
		Set SNAP_ELIG_APPROVALS(snap_elig_months_count) = new snap_eligibility_detail

		SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month = MAXIS_footer_month
		SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_year = MAXIS_footer_year

		Call SNAP_ELIG_APPROVALS(snap_elig_months_count).read_elig

		SNAP_ELIG_APPROVALS(snap_elig_months_count).er_month = False
		SNAP_ELIG_APPROVALS(snap_elig_months_count).hrf_month = False

		If STAT_INFORMATION(month_count).stat_revw_snap_code <> "" Then
			SNAP_ELIG_APPROVALS(snap_elig_months_count).er_month = True
			SNAP_ELIG_APPROVALS(snap_elig_months_count).er_status = STAT_INFORMATION(month_count).stat_revw_snap_code
			SNAP_ELIG_APPROVALS(snap_elig_months_count).er_caf_date = STAT_INFORMATION(month_count).stat_revw_form_recvd_date
			SNAP_ELIG_APPROVALS(snap_elig_months_count).er_interview_date = STAT_INFORMATION(month_count).stat_revw_interview_date
		End If
		If STAT_INFORMATION(month_count).stat_mont_snap_status <> "" Then
			SNAP_ELIG_APPROVALS(snap_elig_months_count).hrf_month = True
			SNAP_ELIG_APPROVALS(snap_elig_months_count).hrf_status = STAT_INFORMATION(month_count).stat_mont_snap_status
			SNAP_ELIG_APPROVALS(snap_elig_months_count).hrf_doc_date = STAT_INFORMATION(month_count).stat_mont_form_recvd_date
		End If

		If first_SNAP_approval = "" AND SNAP_ELIG_APPROVALS(snap_elig_months_count).approved_today = True Then first_SNAP_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
		' MsgBox "SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month - " & SNAP_ELIG_APPROVALS(snap_elig_months_count).elig_footer_month
		SNAP_ELIG_APPROVALS(snap_elig_months_count).adults_recv_snap = 0
		SNAP_ELIG_APPROVALS(snap_elig_months_count).children_recv_snap = 0
		For each_elig_memb = 0 to UBound(SNAP_ELIG_APPROVALS(snap_elig_months_count).snap_elig_ref_numbs)
			For each_stat_memb = 0 to UBound(STAT_INFORMATION(month_count).stat_memb_ref_numb)
				If SNAP_ELIG_APPROVALS(snap_elig_months_count).snap_elig_ref_numbs(each_elig_memb) = STAT_INFORMATION(month_count).stat_memb_ref_numb(each_stat_memb) Then
					If SNAP_ELIG_APPROVALS(snap_elig_months_count).snap_elig_membs_counted(each_elig_memb) <> "COUNTED" Then
						STAT_INFORMATION(month_count).stat_jobs_one_job_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_jobs_two_job_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_jobs_three_job_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_jobs_four_job_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_jobs_five_job_counted(each_stat_memb) = False

						STAT_INFORMATION(month_count).stat_busi_one_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_busi_two_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_busi_three_counted(each_stat_memb) = False

						STAT_INFORMATION(month_count).stat_unea_one_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_unea_two_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_unea_three_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_unea_four_counted(each_stat_memb) = False
						STAT_INFORMATION(month_count).stat_unea_five_counted(each_stat_memb) = False
					End If
					If SNAP_ELIG_APPROVALS(snap_elig_months_count).snap_elig_membs_eligibility(each_elig_memb) = "ELIGIBLE" Then
						If STAT_INFORMATION(month_count).stat_memb_age(each_stat_memb) > 21 Then
							SNAP_ELIG_APPROVALS(snap_elig_months_count).adults_recv_snap = SNAP_ELIG_APPROVALS(snap_elig_months_count).adults_recv_snap + 1
						Else
							SNAP_ELIG_APPROVALS(snap_elig_months_count).children_recv_snap = SNAP_ELIG_APPROVALS(snap_elig_months_count).children_recv_snap + 1
						End If
					End If
				End If
			Next
		Next

		snap_elig_months_count = snap_elig_months_count + 1
	End If


	reDim preserve HC_ELIG_APPROVALS(hc_elig_months_count)

	Set HC_ELIG_APPROVALS(hc_elig_months_count) = new hc_eligibility_detail

	HC_ELIG_APPROVALS(hc_elig_months_count).elig_footer_month = MAXIS_footer_month
	HC_ELIG_APPROVALS(hc_elig_months_count).elig_footer_year = MAXIS_footer_year

	Call HC_ELIG_APPROVALS(hc_elig_months_count).read_elig

	HC_ELIG_APPROVALS(hc_elig_months_count).er_month = False
	HC_ELIG_APPROVALS(hc_elig_months_count).hrf_month = False

	If STAT_INFORMATION(month_count).stat_revw_hc_code <> "" Then
		HC_ELIG_APPROVALS(hc_elig_months_count).er_month = True
		HC_ELIG_APPROVALS(hc_elig_months_count).er_status = STAT_INFORMATION(month_count).stat_revw_hc_code
		HC_ELIG_APPROVALS(hc_elig_months_count).er_caf_date = STAT_INFORMATION(month_count).stat_revw_form_recvd_date
		HC_ELIG_APPROVALS(hc_elig_months_count).er_interview_date = STAT_INFORMATION(month_count).stat_revw_interview_date
	End If
	If STAT_INFORMATION(month_count).stat_mont_hc_status <> "" Then
		HC_ELIG_APPROVALS(hc_elig_months_count).hrf_month = True
		HC_ELIG_APPROVALS(hc_elig_months_count).hrf_status = STAT_INFORMATION(month_count).stat_mont_hc_status
		HC_ELIG_APPROVALS(hc_elig_months_count).hrf_doc_date = STAT_INFORMATION(month_count).stat_mont_form_recvd_date
	End If

	If first_HC_approval = "" AND HC_ELIG_APPROVALS(hc_elig_months_count).approved_today Then first_HC_approval = MAXIS_footer_month & "/" & MAXIS_footer_year
	' elig_list = ""
	' for hc_elig = 0 to UBound(HC_ELIG_APPROVALS(hc_elig_months_count).hc_elig_ref_numbs)
	' 	If HC_ELIG_APPROVALS(hc_elig_months_count).hc_prog_elig_appd(hc_elig) = True Then
	' 		elig_list = elig_list & "MEMB " & HC_ELIG_APPROVALS(hc_elig_months_count).hc_elig_ref_numbs(hc_elig) & ": " & HC_ELIG_APPROVALS(hc_elig_months_count).hc_prog_elig_major_program(hc_elig) & " " & HC_ELIG_APPROVALS(hc_elig_months_count).hc_prog_elig_elig_type(hc_elig) & "-" & HC_ELIG_APPROVALS(hc_elig_months_count).hc_prog_elig_elig_standard(hc_elig) & vbCr
	' 	End If
	' 	' MsgBox "hc_elig - " & hc_elig & vbCr & elig_list
	' Next
	'
	' MsgBox "Footer Month - " & HC_ELIG_APPROVALS(hc_elig_months_count).elig_footer_month & "/" & HC_ELIG_APPROVALS(hc_elig_months_count).elig_footer_year & vbCr &_
	' 	   "HC Eligibility: " & vbCr &_
	' 	   elig_list

	hc_elig_months_count = hc_elig_months_count + 1


	month_count = month_count + 1					'This is way down here because I want to be able to reference the information in the current month for this class.


	Call back_to_SELF
Next

EMWriteScreen MAXIS_case_number, 18, 43


pnd2_display_limit_hit = False
deny_app_one = False
deny_app_two = False
denials_found_on_pnd2 = False
advise_not_to_use_lcase = False
Call navigate_to_MAXIS_screen("REPT", "PND2")
EMReadScreen pnd2_disp_limit, 13, 6, 35             'functionality to bypass the display limit warning if it appears.
If pnd2_disp_limit = "Display Limit" Then
	TRANSMIT
	pnd2_display_limit_hit = True
End If

row = 1
col = 1
EMSearch MAXIS_case_number, row, col
If row <> 24 and row <> 0 Then
	EMReadScreen pnd2_appl_date, 8, row, 38
	EMReadScreen pnd2_days_pending, 5, row, 48
	EMReadScreen pnd2_cash_status, 1, row, 54
	EMReadScreen pnd2_cash_prog_one, 2, row, 56
	EMReadScreen pnd2_cash_prog_two, 2, row, 59
	EMReadScreen pnd2_snap_status, 1, row, 62
	' EMReadScreen pnd2_hc_status, 1, row,  65
	EMReadScreen pnd2_emer_status, 1, row, 68
	EMReadScreen pnd2_grh_status, 1, row, 72
	' EMReadScreen pnd2_ive_status, 1, row, 76
	pnd2_days_pending = trim(pnd2_days_pending)

	If pnd2_cash_status = "i" Then EMWriteScreen "I", row, 54
	If pnd2_cash_status = "r" Then EMWriteScreen "R", row, 54
	If pnd2_snap_status = "i" Then EMWriteScreen "I", row, 62
	If pnd2_snap_status = "r" Then EMWriteScreen "R", row, 62
	If pnd2_emer_status = "i" Then EMWriteScreen "I", row, 68
	If pnd2_emer_status = "r" Then EMWriteScreen "R", row, 68
	If pnd2_grh_status = "i" Then EMWriteScreen "I", row, 72
	If pnd2_grh_status = "r" Then EMWriteScreen "R", row, 72

	If pnd2_cash_status = "i" or pnd2_cash_status = "r" or pnd2_snap_status = "i" or pnd2_snap_status = "r" or pnd2_emer_status = "i" or pnd2_emer_status = "r" or pnd2_grh_status = "i" or pnd2_grh_status = "r" Then advise_not_to_use_lcase = True

	pnd2_cash_status = UCase(pnd2_cash_status)
	pnd2_snap_status = UCase(pnd2_snap_status)
	pnd2_emer_status = UCase(pnd2_emer_status)
	pnd2_grh_status = UCase(pnd2_grh_status)


	If pnd2_cash_status = "I" Then deny_app_one = True
	If pnd2_cash_status = "R" Then deny_app_one = True
	If pnd2_snap_status = "I" Then deny_app_one = True
	If pnd2_snap_status = "R" Then deny_app_one = True
	If pnd2_emer_status = "I" Then deny_app_one = True
	If pnd2_emer_status = "R" Then deny_app_one = True
	If pnd2_grh_status = "I" Then deny_app_one = True
	If pnd2_grh_status = "R" Then deny_app_one = True

	If pnd2_cash_prog_one = "  " Then pnd2_cash_prog_one = ""
	If pnd2_cash_prog_one = "CA" Then pnd2_cash_prog_one = "Cash"
	If pnd2_cash_prog_one = "MF" Then pnd2_cash_prog_one = "MFIP"
	If pnd2_cash_prog_one = "DW" Then pnd2_cash_prog_one = "DWP"
	If pnd2_cash_prog_one = "MS" Then pnd2_cash_prog_one = "MSA"
	If pnd2_cash_prog_one = "RC" Then pnd2_cash_prog_one = "RCA"
	If pnd2_cash_prog_two = "  " Then pnd2_cash_prog_two = ""
	If pnd2_cash_prog_two = "CA" Then pnd2_cash_prog_two = "Cash"
	If pnd2_cash_prog_two = "MF" Then pnd2_cash_prog_two = "MFIP"
	If pnd2_cash_prog_two = "DW" Then pnd2_cash_prog_two = "DWP"
	If pnd2_cash_prog_two = "MS" Then pnd2_cash_prog_two = "MSA"
	If pnd2_cash_prog_two = "RC" Then pnd2_cash_prog_two = "RCA"

	pnd2_appl_date = replace(pnd2_appl_date, " ", "/")

	row = row + 1
	EMReadScreen additional_application_check, 14, row, 17                 'looking to see if this case has a secondary application date entered
	IF additional_application_check = "ADDITIONAL APP" THEN                         'If it does this string will be at that location and we need to do some handling around the application date to use.
		EMReadScreen pnd2_2nd_appl_date, 8, row, 38
		EMReadScreen pnd2_2nd_days_pending, 5, row, 48
		EMReadScreen pnd2_2nd_cash_status, 1, row, 54
		EMReadScreen pnd2_2nd_cash_prog_one, 2, row, 56
		EMReadScreen pnd2_2nd_cash_prog_two, 2, row, 59
		EMReadScreen pnd2_2nd_snap_status, 1, row, 62
		' EMReadScreen pnd2_hc_status, 1, row,  65
		EMReadScreen pnd2_2nd_emer_status, 1, row, 68
		EMReadScreen pnd2_2nd_grh_status, 1, row, 72
		' EMReadScreen pnd2_ive_status, 1, row, 76
		pnd2_days_pending = trim(pnd2_days_pending)

		If pnd2_2nd_cash_status = "i" Then EMWriteScreen "I", row, 54
		If pnd2_2nd_cash_status = "r" Then EMWriteScreen "R", row, 54
		If pnd2_2nd_snap_status = "i" Then EMWriteScreen "I", row, 62
		If pnd2_2nd_snap_status = "r" Then EMWriteScreen "R", row, 62
		If pnd2_2nd_emer_status = "i" Then EMWriteScreen "I", row, 68
		If pnd2_2nd_emer_status = "r" Then EMWriteScreen "R", row, 68
		If pnd2_2nd_grh_status = "i" Then EMWriteScreen "I", row, 72
		If pnd2_2nd_grh_status = "r" Then EMWriteScreen "R", row, 72

		If pnd2_2nd_cash_status = "i" or pnd2_2nd_cash_status = "r" or pnd2_2nd_snap_status = "i" or pnd2_2nd_snap_status = "r" or pnd2_2nd_emer_status = "i" or pnd2_2nd_emer_status = "r" or pnd2_2nd_grh_status = "i" or pnd2_2nd_grh_status = "r" Then advise_not_to_use_lcase = True

		pnd2_2nd_cash_status = UCase(pnd2_2nd_cash_status)
		pnd2_2nd_snap_status = UCase(pnd2_2nd_snap_status)
		pnd2_2nd_emer_status = UCase(pnd2_2nd_emer_status)
		pnd2_2nd_grh_status = UCase(pnd2_2nd_grh_status)

		' If pnd2_2nd_grh_status = "_" Then pnd2_2nd_grh_status = ""
		' If pnd2_2nd_grh_status = "_" Then pnd2_2nd_grh_status = ""

		If pnd2_2nd_cash_status = "I" Then deny_app_two = True
		If pnd2_2nd_cash_status = "R" Then deny_app_two = True
		If pnd2_2nd_snap_status = "I" Then deny_app_two = True
		If pnd2_2nd_snap_status = "R" Then deny_app_two = True
		If pnd2_2nd_emer_status = "I" Then deny_app_two = True
		If pnd2_2nd_emer_status = "R" Then deny_app_two = True
		If pnd2_2nd_grh_status = "I" Then deny_app_two = True
		If pnd2_2nd_grh_status = "R" Then deny_app_two = True

		If pnd2_2nd_cash_prog_one = "  " Then pnd2_2nd_cash_prog_one = ""
		If pnd2_2nd_cash_prog_one = "CA" Then pnd2_2nd_cash_prog_one = "Cash"
		If pnd2_2nd_cash_prog_one = "MF" Then pnd2_2nd_cash_prog_one = "MFIP"
		If pnd2_2nd_cash_prog_one = "DW" Then pnd2_2nd_cash_prog_one = "DWP"
		If pnd2_2nd_cash_prog_one = "MS" Then pnd2_2nd_cash_prog_one = "MSA"
		If pnd2_2nd_cash_prog_one = "RC" Then pnd2_2nd_cash_prog_one = "RCA"
		If pnd2_2nd_cash_prog_two = "  " Then pnd2_2nd_cash_prog_two = ""
		If pnd2_2nd_cash_prog_two = "CA" Then pnd2_2nd_cash_prog_two = "Cash"
		If pnd2_2nd_cash_prog_two = "MF" Then pnd2_2nd_cash_prog_two = "MFIP"
		If pnd2_2nd_cash_prog_two = "DW" Then pnd2_2nd_cash_prog_two = "DWP"
		If pnd2_2nd_cash_prog_two = "MS" Then pnd2_2nd_cash_prog_two = "MSA"
		If pnd2_2nd_cash_prog_two = "RC" Then pnd2_2nd_cash_prog_two = "RCA"

		pnd2_2nd_appl_date = replace(pnd2_2nd_appl_date, " ", "/")
	End If

	If deny_app_one = True or deny_app_two = True Then denials_found_on_pnd2 = True
	If advise_not_to_use_lcase = True Then MsgBox "The script has reviewed REPT/PND2 and found that denial coding has been entered as lowercase letters. " & vbCr & vbCr & " - Either an 'i' or 'r' was entered for the case to deny." & vbCr & vbCr & "Please be aware that entering command information in MAXIS is most reliable if entered as an Upper Case letter." & vbCr & "The script has repaired the entry to make them capitalized." & vbCr & vbCr & "This information is just for awareness, the script will continue and there is no additional action needed."
End If

Call back_to_SELF
EMWriteScreen MAXIS_case_number, 18, 43


'In order to determine the array - need to be able to see if the budget changes from one to the next
'EMER doesn't have an array - there is only one month

If first_DWP_approval <> "" Then enter_CNOTE_for_DWP = True
If first_MFIP_approval <> "" Then enter_CNOTE_for_MFIP = True
If first_MSA_approval <> "" Then enter_CNOTE_for_MSA = True
If first_GA_approval <> "" Then enter_CNOTE_for_GA = True
If first_DENY_approval <> "" Then enter_CNOTE_for_DENY = True
If first_GRH_approval <> "" Then enter_CNOTE_for_GRH = True
If first_SNAP_approval <> "" Then enter_CNOTE_for_SNAP = True
If first_HC_approval <> "" Then enter_CNOTE_for_HC = True
' MsgBox "first_SNAP_approval - " & first_SNAP_approval & vbCr & "enter_CNOTE_for_SNAP - " & enter_CNOTE_for_SNAP

deductions_detail_btn 	= 10
hh_comp_detail			= 20
shel_exp_detail_btn		= 30
unique_approval_explain_btn	= 40
nav_stat_elig_btn		= 50

app_confirmed_btn		= 100
next_approval_btn		= 110
app_incorrect_btn		= 120

const months_in_approval			= 0
' const wcom_needed 					= 4
const verif_reqquest_date			= 5
const pact_inelig_reasons			= 6
const package_is_expedited_const 	= 7
const include_budget_in_note_const	= 8
const confirm_budget_selection		= 9
const first_mo_const				= 10
const last_mo_const					= 11
const wcom_needed					= 12
const dialog_displayed				= 13
const budget_error_exists			= 14
const btn_one 						= 15
const approval_incorrect			= 16
const pact_wcom_needed				= 17
const pact_wcom_sent				= 18
const snap_over_130_wcom_needed		= 19
const snap_over_130_wcom_sent		= 20
const snap_130_percent_fpg_amt		= 21

const approval_confirmed			= 22

Dim SNAP_UNIQUE_APPROVALS()
ReDim SNAP_UNIQUE_APPROVALS(approval_confirmed, 0)

'TODO - Add functionality to review CASE/NOTES for approvals that have been noted TODAY so that we don't double up on NOTES
If enter_CNOTE_for_SNAP = True Then
	'Budgets match if the earned income, unearned income, shelter expenses, and entitlements are the same

	last_earned_income = ""
	last_unearned_income = ""
	last_shelter_expense = ""
	last_hest_expense = ""
	last_snap_entitlement = ""
	start_capturing_approvals = False
	unique_app_count = 0
	For approval = 0 to UBound(SNAP_ELIG_APPROVALS)

		' SNAP_ELIG_APPROVALS(approval).gather_stat_info

		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_total_earned_inc)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_total_unea_inc)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_rent_mort)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_prop_tax)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_home_ins)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_other)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_electricity)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_heat_ac)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_budg_shel_phone)
		' Call ensure_variable_is_a_number(SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot)

		' sum_housing = SNAP_ELIG_APPROVALS(approval).snap_budg_shel_rent_mort + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_prop_tax + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_home_ins + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_other
		' sum_hest = SNAP_ELIG_APPROVALS(approval).snap_budg_shel_electricity + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_heat_ac + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_phone
		'
		' MsgBox ("Earned Income - " & SNAP_ELIG_APPROVALS(approval).snap_budg_total_earned_inc & vbCr & "Unearned Income - " & SNAP_ELIG_APPROVALS(approval).snap_budg_total_unea_inc & vbCr & "Housing Exp - " & sum_housing & vbCr & "Hest - " & sum_hest & vbCr & "FS Allotment - " & SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot & vbCr & vbCr & "Month - " & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year)
		If SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year = first_SNAP_approval Then start_capturing_approvals = True
		If start_capturing_approvals = True Then
			If unique_app_count = 0 Then
				ReDim preserve SNAP_UNIQUE_APPROVALS(approval_confirmed, unique_app_count)

				SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app_count) = SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year
				SNAP_UNIQUE_APPROVALS(first_mo_const, unique_app_count) = SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year
				SNAP_UNIQUE_APPROVALS(btn_one, unique_app_count) = 500 + unique_app_count
				SNAP_UNIQUE_APPROVALS(approval_confirmed, unique_app_count) = False
				SNAP_UNIQUE_APPROVALS(approval_incorrect, unique_app_count) = False
				SNAP_UNIQUE_APPROVALS(package_is_expedited_const, unique_app_count) = SNAP_ELIG_APPROVALS(approval).snap_expedited
				SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app_count) = True
				last_earned_income = SNAP_ELIG_APPROVALS(approval).snap_budg_total_earned_inc
				last_unearned_income = SNAP_ELIG_APPROVALS(approval).snap_budg_total_unea_inc
				last_shelter_expense = SNAP_ELIG_APPROVALS(approval).snap_budg_shel_rent_mort + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_prop_tax + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_home_ins + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_other
				last_hest_expense = SNAP_ELIG_APPROVALS(approval).snap_budg_shel_electricity + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_heat_ac + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_phone
				last_eligibility = SNAP_ELIG_APPROVALS(approval).snap_elig_result
				last_snap_entitlement = SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot
				last_expedited_status = SNAP_ELIG_APPROVALS(approval).snap_expedited

				unique_app_count = unique_app_count + 1
			Else
				match_last_benefit_amounts = True

				If last_earned_income <> SNAP_ELIG_APPROVALS(approval).snap_budg_total_earned_inc Then match_last_benefit_amounts = False
				If last_unearned_income <> SNAP_ELIG_APPROVALS(approval).snap_budg_total_unea_inc Then match_last_benefit_amounts = False
				If last_shelter_expense <> SNAP_ELIG_APPROVALS(approval).snap_budg_shel_rent_mort + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_prop_tax + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_home_ins + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_other Then match_last_benefit_amounts = False
				If last_hest_expense <> SNAP_ELIG_APPROVALS(approval).snap_budg_shel_electricity + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_heat_ac + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_phone Then match_last_benefit_amounts = False
				If last_eligibility <> SNAP_ELIG_APPROVALS(approval).snap_elig_result Then match_last_benefit_amounts = False
				If last_snap_entitlement <> SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot Then match_last_benefit_amounts = False
				If last_expedited_status <> SNAP_ELIG_APPROVALS(approval).snap_expedited Then match_last_benefit_amounts = False

				If match_last_benefit_amounts = True Then
					SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app_count-1) = SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app_count-1) & "~" & SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year
					SNAP_UNIQUE_APPROVALS(last_mo_const, unique_app_count-1) = SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year
				End If

				If match_last_benefit_amounts = False Then
					ReDim preserve SNAP_UNIQUE_APPROVALS(approval_confirmed, unique_app_count)

					SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app_count) = SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year
					SNAP_UNIQUE_APPROVALS(first_mo_const, unique_app_count) = SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year
					SNAP_UNIQUE_APPROVALS(btn_one, unique_app_count) = 500 + unique_app_count
					SNAP_UNIQUE_APPROVALS(approval_confirmed, unique_app_count) = False
					SNAP_UNIQUE_APPROVALS(approval_incorrect, unique_app_count) = False
					SNAP_UNIQUE_APPROVALS(package_is_expedited_const, unique_app_count) = SNAP_ELIG_APPROVALS(approval).snap_expedited
					SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app_count) = True
					last_earned_income = SNAP_ELIG_APPROVALS(approval).snap_budg_total_earned_inc
					last_unearned_income = SNAP_ELIG_APPROVALS(approval).snap_budg_total_unea_inc
					last_shelter_expense = SNAP_ELIG_APPROVALS(approval).snap_budg_shel_rent_mort + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_prop_tax + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_home_ins + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_other
					last_hest_expense = SNAP_ELIG_APPROVALS(approval).snap_budg_shel_electricity + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_heat_ac + SNAP_ELIG_APPROVALS(approval).snap_budg_shel_phone
					last_eligibility = SNAP_ELIG_APPROVALS(approval).snap_elig_result
					last_snap_entitlement = SNAP_ELIG_APPROVALS(approval).snap_benefit_monthly_fs_allot
					last_expedited_status = SNAP_ELIG_APPROVALS(approval).snap_expedited

					' MsgBox ("last_shelter_expense - " & last_shelter_expense & vbCr & "last_hest_expense - " & last_hest_expense)

					unique_app_count = unique_app_count + 1
				End If
			End If
		End If
	Next

	' For unique_app = 0 to UBound(SNAP_UNIQUE_APPROVALS, 2)
	' Next

	all_snap_approvals_confirmed = False
	approval_selected = 0

	Do

		first_month = left(SNAP_UNIQUE_APPROVALS(months_in_approval, approval_selected), 5)
		elig_ind = ""
		month_ind = ""
		For approval = 0 to UBound(SNAP_ELIG_APPROVALS)
			If SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year = first_month Then elig_ind = approval
		Next
		For each_month = 0 to UBound(STAT_INFORMATION)
			If STAT_INFORMATION(each_month).footer_month & "/" & STAT_INFORMATION(each_month).footer_year = first_month Then month_ind = each_month
		Next

		If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "INELIGIBLE" Then
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_appl_withdrawn_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_applct_elig_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_comdty_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_disq_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_dupl_assist_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_eligible_person_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_fail_coop_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_fail_file_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			' snap_case_prosp_gross_inc_test
			' snap_case_prosp_net_inc_test

			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_recert_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_residence_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_resource_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			' snap_case_retro_gross_inc_test
			' snap_case_retro_net_inc_test
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_strike_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_xfer_resource_inc_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_voltry_quit_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_work_reg_test = "FAILED" Then SNAP_UNIQUE_APPROVALS(include_budget_in_note_const, unique_app) = False
		End If

		Call determine_130_percent_of_FPG(left(first_month, 2), right(first_month, 2), SNAP_ELIG_APPROVALS(elig_ind).snap_budg_numb_in_assist_unit, SNAP_UNIQUE_APPROVALS(snap_130_percent_fpg_amt, unique_app))
		SNAP_UNIQUE_APPROVALS(pact_wcom_needed, unique_app) = False
		SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_needed, unique_app) = False
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_PACT = "FAILED" then SNAP_UNIQUE_APPROVALS(pact_wcom_needed, unique_app) = True


		If IsNumeric(SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_gross_inc) = True and IsNumeric(SNAP_UNIQUE_APPROVALS(snap_130_percent_fpg_amt, unique_app)) = True Then
			grs_inc = SNAP_ELIG_APPROVALS(elig_ind).snap_budg_total_gross_inc*1
			max_130_inc = SNAP_UNIQUE_APPROVALS(snap_130_percent_fpg_amt, unique_app)*1
			If grs_inc > max_130_inc AND SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "ELIGIBLE" Then SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_needed, unique_app) = True
		End If
		SNAP_UNIQUE_APPROVALS(pact_wcom_sent, unique_app) = False
		SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_sent, unique_app) = False

		SNAP_UNIQUE_APPROVALS(wcom_needed, unique_app) = False
		If SNAP_UNIQUE_APPROVALS(pact_wcom_needed, unique_app) = True Then SNAP_UNIQUE_APPROVALS(wcom_needed, unique_app) = True
		If SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_needed, unique_app) = True Then SNAP_UNIQUE_APPROVALS(wcom_needed, unique_app) = True

		ei_count = 0
		unea_count = 0
		For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
		  If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_one_job_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_two_job_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_three_job_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_four_job_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_jobs_five_job_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_one_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_two_counted(each_memb) = True Then ei_count = ei_count + 1
		  If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_busi_three_counted(each_memb) = True Then ei_count = ei_count + 1

		  If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_one_counted(each_memb) = True Then
			  unea_count = unea_count + 1
			  If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then unea_count = unea_count + 1
		  End If
		  If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_two_counted(each_memb) = True Then
			  unea_count = unea_count + 1
			  If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then unea_count = unea_count + 1
		  End If
		  If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_three_counted(each_memb) = True Then
			  unea_count = unea_count + 1
			  If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then unea_count = unea_count + 1
		  End If
		  If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_four_counted(each_memb) = True Then
			  unea_count = unea_count + 1
			  If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then unea_count = unea_count + 1
		  End If
		  If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True AND STAT_INFORMATION(month_ind).stat_unea_five_counted(each_memb) = True Then
			  unea_count = unea_count + 1
			  If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then unea_count = unea_count + 1
		  End If
		Next
		ei_len = ei_count * 20
		unea_len = unea_count * 10
		income_box_len = 30 + unea_len
		If ei_len > unea_count Then income_box_len = 30 + ei_len

		call snap_elig_dialog

		dialog Dialog1
		cancel_confirmation

		err_msg = ""


		If right(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected), 1) = "." Then SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected) = left(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected), len(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected))- 1)
		If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test = "FAILED" and SNAP_UNIQUE_APPROVALS(confirm_budget_selection, approval_selected) <> "No - I need to complete a new Approval" then
			If Isdate(SNAP_UNIQUE_APPROVALS(verif_reqquest_date, approval_selected)) = False Then
				err_msg = err_msg & vbNewLine & "* Enter the date the verification request form sent from ECF to detail information about missing verifications for an Ineligible SNAP approval."
			Else
				If DateDiff("d", SNAP_UNIQUE_APPROVALS(verif_reqquest_date, approval_selected), date) < 10 AND SNAP_UNIQUE_APPROVALS(confirm_budget_selection, approval_selected) = "Yes - budget is Accurate" Then
					err_msg = err_msg & vbNewLine & "* The verification request date: " &  SNAP_UNIQUE_APPROVALS(verif_reqquest_date, approval_selected) & " is less than 10 days ago and we should not be taking action yet."
					SNAP_UNIQUE_APPROVALS(confirm_budget_selection, approval_selected) = "No - I need to complete a new Approval"
				End If
			End If
			If SNAP_ELIG_APPROVALS(elig_ind).snap_case_verif_test_PACT = "FAILED" then
				If trim(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected)) = "" Then
					err_msg = err_msg & vbNewLine & "* Since PACT was used to approve this SNAP benefit as ineligible, list the reasons for ineligibility."
				ElseIf len(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected)) < 30 Then
					err_msg = err_msg & vbNewLine & "* SNAP ineligibility due to PACT requires sufficient explaination, expand upon the information entered in the Reason for Ineligibility field."
				End If
				If trim(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected)) = "" or len(SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, approval_selected)) < 15 Then err_msg = err_msg & vbNewLine & " *** This information will be entered in a WCOM and should be writen without appreviations and in full detail."
			End if
		End if

		If err_msg <> "" and ButtonPressed < 1100 Then
			MsgBox "*** INFORMATION IN SCRIPT DIALOG INCOMPLETE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
			If ButtonPressed = app_confirmed_btn Then ButtonPressed = -1
		End If

		If ButtonPressed = nav_stat_elig_btn Then
			ft_mo = left(first_month, 2)
			ft_yr = right(first_month, 2)
			Call back_to_SELF
			call navigate_to_MAXIS_screen("ELIG", "FS  ")
			EMWriteScreen ft_mo, 19, 54
			EMWriteScreen ft_yr, 19, 57
			Call find_last_approved_ELIG_version(19, 78, vrs_numb, vrs_dt, vrs_rslt, approval_found)
			' transmit
		End If

		If ButtonPressed = deductions_detail_btn then MsgBox "DEDUCTION EXPLANATION TO GO HERE"
		If ButtonPressed = hh_comp_detail then MsgBox "HH COMP EXPLANATION TO GO HERE"
		If ButtonPressed = shel_exp_detail_btn then MsgBox "SHELTER EXPENSE EXPLANATION TO GO HERE"
		If ButtonPressed = unique_approval_explain_btn then MsgBox "UNIQUE APPROVAL EXPLANATION TO GO HERE"
		' If ButtonPressed = app_confirmed_btn

		If err_msg = "" Then

			all_snap_approvals_confirmed = True
			snap_approval_is_incorrect = False

			If SNAP_UNIQUE_APPROVALS(confirm_budget_selection, approval_selected) = "Yes - budget is Accurate" Then
				SNAP_UNIQUE_APPROVALS(approval_confirmed, approval_selected) = True
				SNAP_UNIQUE_APPROVALS(approval_incorrect, approval_selected) = False
			ElseIf SNAP_UNIQUE_APPROVALS(confirm_budget_selection, approval_selected) = "No - I need to complete a new Approval" Then
				SNAP_UNIQUE_APPROVALS(approval_confirmed, approval_selected) = False
				SNAP_UNIQUE_APPROVALS(approval_incorrect, approval_selected) = True
			End If

			If ButtonPressed = -1 Then
				If approval_selected = UBound(SNAP_UNIQUE_APPROVALS, 2) Then
					ButtonPressed = app_confirmed_btn
				ElseIf snap_approval_is_incorrect = True Then
					ButtonPressed = app_incorrect_btn
				Else
					ButtonPressed = next_approval_btn
				End If
			End If

			not_confirmed_pckg_list = ""
			first_unconfirmmed_month = ""
			for each_app = 0 to UBound(SNAP_UNIQUE_APPROVALS, 2)
				If ButtonPressed = SNAP_UNIQUE_APPROVALS(btn_one, each_app) Then approval_selected = each_app
				If SNAP_UNIQUE_APPROVALS(approval_confirmed, each_app) = False Then
					all_snap_approvals_confirmed = False
					not_confirmed_pckg_list = not_confirmed_pckg_list & replace(SNAP_UNIQUE_APPROVALS(months_in_approval, each_app), "~", " - ") & vbCr
					If first_unconfirmmed_month = "" Then first_unconfirmmed_month = each_app
				End If
				If SNAP_UNIQUE_APPROVALS(approval_incorrect, each_app) = True Then snap_approval_is_incorrect = True
			Next

			If ButtonPressed = next_approval_btn Then
				approval_selected = approval_selected + 1
				If approval_selected > UBound(SNAP_UNIQUE_APPROVALS, 2) Then
					If all_snap_approvals_confirmed = True Then
						ButtonPressed = app_confirmed_btn
					Else
						approval_selected = UBound(SNAP_UNIQUE_APPROVALS, 2)
					End If
				End If
			End If
		End If
		If ButtonPressed = app_confirmed_btn and all_snap_approvals_confirmed = False Then
			MsgBox "*** All Approval Packages need to be Confirmed ****" & vbCr & vbCr & "Please review all the approval packages and indicate if they are correct before the scrript can continue." & vbCr & vbCr & "Review the following approval package(s)" & vbCr & not_confirmed_pckg_list
			approval_selected = first_unconfirmmed_month
		End If
		' For unique_app = 0 to UBound(SNAP_UNIQUE_APPROVALS, 2)
		' Next

	Loop until (ButtonPressed = app_confirmed_btn and all_snap_approvals_confirmed = True) or ButtonPressed = app_incorrect_btn

	If snap_approval_is_incorrect = True Then
		enter_CNOTE_for_SNAP = False
		end_msg_info = end_msg_info & "CASE/NOTE has NOT been entered for SNAP Approvals from " & first_SNAP_approval & " onward as the approval appears incorrect and needs to be updated and ReApproved." & vbCr
	End if

End If


If enter_CNOTE_for_SNAP = True Then
	' MsgBox "MADE IT TO THE NOTE"
	''
	For unique_app = 0 to UBound(SNAP_UNIQUE_APPROVALS, 2)
		first_month = left(SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app), 5)
		If len(SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app)) > 5 Then
			last_month = right(SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app), 5)
		End If

		elig_ind = ""
		one_month_is_elig = False
		For approval = 0 to UBound(SNAP_ELIG_APPROVALS)
			If SNAP_ELIG_APPROVALS(approval).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval).elig_footer_year = first_month Then elig_ind = approval
		Next
		month_ind = ""
		For each_month = 0 to UBound(STAT_INFORMATION)
			If STAT_INFORMATION(each_month).footer_month & "/" & STAT_INFORMATION(each_month).footer_year = first_month Then month_ind = each_month
		Next

		program_detail = "- SNAP"
		header_end = ""
		If SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "ELIGIBLE" Then
			If last_month = curr_month_plus_one or first_month = curr_month_plus_one Then
				header_end = " - Ongoing"
			ElseIf len(SNAP_UNIQUE_APPROVALS(months_in_approval, unique_app)) > 5 Then
				header_end = " - " & last_month
			Else
				header_end = " only"
			End If
			If SNAP_UNIQUE_APPROVALS(package_is_expedited_const, unique_app) = True Then program_detail = "EXPEDITED SNAP"
			elig_info = "ELIGIBLE"
			one_month_is_elig = True
		ElseIf SNAP_ELIG_APPROVALS(elig_ind).snap_elig_result = "INELIGIBLE" Then
			If snap_status = "INACTIVE" Then elig_info = "INELIGIBLE - Denied"
			If snap_status = "APP OPEN" Then elig_info = "INELIGIBLE - Denied"
			If snap_status = "APP CLOSE" Then elig_info = "INELIGIBLE - Closed"
			If one_month_is_elig = True Then elig_info = "INELIGIBLE - Closed"
		End If
		due_date = ""
		If IsDate(SNAP_UNIQUE_APPROVALS(verif_reqquest_date, approval_selected)) = True Then due_date = DateAdd("d", 10, SNAP_UNIQUE_APPROVALS(verif_reqquest_date, unique_app))

		'This is the WCOM part
		If SNAP_UNIQUE_APPROVALS(wcom_needed, unique_app) = True Then
			ft_mo = left(first_month, 2)
			ft_yr = right(first_month, 2)

			Call navigate_to_MAXIS_screen("SPEC", "WCOM")
			EMWriteScreen ft_mo, 03, 46
			EMWriteScreen ft_yr, 03, 51
			transmit

			wcom_row = 7
			Do
				EMReadScreen notc_date, 8, wcom_row, 16
				EMReadScreen notc_type, 2, wcom_row, 26
				EMReadScreen notc_description, 30, wcom_row, 30
				EMReadScreen notc_print_status, 8, wcom_row, 71

				If notc_date <> "        " Then
					notc_date = DateAdd("d", 0, notc_date)
					notc_description = trim(notc_description)
					notc_print_status = trim(notc_print_status)
					If DateDiff("d", date, notc_date) = 0 AND notc_type = "FS" AND notc_description = "ELIG Approval Notice" AND notc_print_status = "Waiting" Then
						Call write_value_and_transmit("X", wcom_row, 13)

						PF9
						EMReadScreen wcom_line, 60, 3, 15
						If trim(wcom_line) = "" Then

							If SNAP_UNIQUE_APPROVALS(pact_wcom_needed, unique_app) = True Then
								If right(elig_info, 6) = "Denied" Then
									' 60_days_from_app = ""
									' If IsDate(STAT_INFORMATION(month_ind).stat_prog_snap_appl_date) = True Then 60_days_from_app = DateAdd("d", 60, STAT_INFORMATION(month_ind).stat_prog_snap_appl_date)
									' "Your SNAP application has been denied because you did not provide: " & SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, unique_app) & ".  This proof was needed by " & due_date & ".  If you need assistance getting this proof please contact us at the number listed on this notice by " & 60_days_from_app"." ''(This date will be 60 days after the application date).
									CALL write_variable_in_SPEC_MEMO("Your SNAP application has been denied because you did not provide: " & SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, unique_app) & ".  This proof was needed by " & due_date & ".  If you need assistance getting this proof please contact us at the number listed on this notice by " & DateAdd("d", 30, date) & ".") ''(This date will be 30 days from today).
								End If

								If right(elig_info, 6) = "Closed" Then
									first_of_closure = ft_mo & "/1/" & ft_yr
									first_of_closure = DateAdd("d", 0, first_of_closure)
									end_of_closure_mo = DateAdd("m", 1, first_of_closure)
									end_of_closure_mo = DateAdd("d", -1, end_of_closure_mo)
									CALL write_variable_in_SPEC_MEMO("Your SNAP case will close because you did not provide: " & SNAP_UNIQUE_APPROVALS(pact_inelig_reasons, unique_app) & ".  This proof was needed by " & due_date & ".  If you need assistance getting this proof please contact us at the number listed on this notice by " & end_of_closure_mo & ".")  ''(Enter the last day of the month prior to the effective date of the closing)"
								End If
								SNAP_UNIQUE_APPROVALS(pact_wcom_sent, unique_app) = True
							End if
							If SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_needed, unique_app) = True Then
								CALL write_variable_in_SPEC_MEMO("The monthly gross income for your household is higher than 130% FPG for your household size. This amount is listed above in this notice.  You do not need to report changes in income until your next renewal.  If you have a decrease in income you may be eligible for more benefits.  You may contact the phone number in this letter if this occurs.")
								SNAP_UNIQUE_APPROVALS(snap_over_130_wcom_sent, unique_app) = True
							End If
							PF4
							PF3
						End If
						Exit Do
					End If
				End if
				wcom_row = wcom_row + 1
			Loop until notc_date = "        "
			Call back_to_SELF
		End If

		'Here we entere the CASE NOTE
		Call snap_elig_case_note
	Next

End If

If denials_found_on_pnd2 = True Then

	Do
		Dlg_len = 65

		If pnd2_cash_status = "I" or pnd2_cash_status = "R" then Dlg_len = Dlg_len + 50
		If (pnd2_cash_status = "I" or pnd2_cash_status = "R") and pnd2_cash_prog_two <> "" then Dlg_len = Dlg_len + 50
		If pnd2_2nd_cash_status = "I" or pnd2_2nd_cash_status = "R" then Dlg_len = Dlg_len + 50
		If (pnd2_2nd_cash_status = "I" or pnd2_2nd_cash_status = "R") and pnd2_2nd_cash_prog_two <> "" then Dlg_len = Dlg_len + 50
		If pnd2_snap_status = "I" or pnd2_snap_status = "R" then Dlg_len = Dlg_len + 50
		If pnd2_2nd_snap_status = "I" or pnd2_2nd_snap_status = "R" then Dlg_len = Dlg_len + 50
		If pnd2_emer_status = "I" or pnd2_emer_status = "R" then Dlg_len = Dlg_len + 50
		If pnd2_2nd_emer_status = "I" or pnd2_2nd_emer_status = "R" then Dlg_len = Dlg_len + 50
		If pnd2_grh_status = "I" or pnd2_grh_status = "R" then Dlg_len = Dlg_len + 50
		If pnd2_2nd_grh_status = "I" or pnd2_2nd_grh_status = "R" then Dlg_len = Dlg_len + 50

		y_pos = 25
		cash_listed = 0

		BeginDialog Dialog1, 0, 0, 341, Dlg_len, "Program Denials Via REPT/PND2"
		  Text 15, 10, 320, 10, "This case has been updated to have denials processed through the REPT/PND2 overnight batch."
		  If pnd2_cash_status = "I" or pnd2_cash_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "Cash Denial"
			  If pnd2_cash_status = "I" Then Text 20, y_pos + 15, 155, 10, "Cash Dened for NO INTERVIEW"
			  If pnd2_cash_status = "R" Then Text 20, y_pos + 15, 155, 10, "Cash Application WITHDRAWN"
			  Text 20, y_pos + 25, 155, 10, "Cash Program: " & pnd2_cash_prog_one
			  Text 185, y_pos + 15, 140, 10, "Cash application date: " & pnd2_appl_date
			  Text 185, y_pos + 25, 140, 10, "Cash has been pending for " & pnd2_days_pending & " Days."

			  cash_listed = cash_listed + 1
			  y_pos = y_pos + 50

			  If pnd2_cash_prog_two <> "" Then
				  GroupBox 15, y_pos, 315, 45, "Cash Denial"
				  If pnd2_cash_status = "I" Then Text 20, y_pos + 15, 155, 10, "Cash Dened for NO INTERVIEW"
				  If pnd2_cash_status = "R" Then Text 20, y_pos + 15, 155, 10, "Cash Application WITHDRAWN"
				  Text 20, y_pos + 25, 155, 10, "Cash Program: " & pnd2_cash_prog_two
				  Text 185, y_pos + 15, 140, 10, "Cash application date: " & pnd2_appl_date
				  Text 185, y_pos + 25, 140, 10, "Cash has been pending for " & pnd2_days_pending & " Days."

				  cash_listed = cash_listed + 1
				  y_pos = y_pos + 50
			  End If
		  End If

		  If pnd2_2nd_cash_status = "I" or pnd2_2nd_cash_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "Cash Denial"
			  If pnd2_2nd_cash_status = "I" Then Text 20, y_pos + 15, 155, 10, "Cash Dened for NO INTERVIEW"
			  If pnd2_2nd_cash_status = "R" Then Text 20, y_pos + 15, 155, 10, "Cash Application WITHDRAWN"
			  Text 20, y_pos + 25, 155, 10, "Cash Program: " & pnd2_2nd_cash_prog_one
			  Text 185, y_pos + 15, 140, 10, "Cash application date: " & pnd2_2nd_appl_date
			  Text 185, y_pos + 25, 140, 10, "Cash has been pending for " & pnd2_2nd_days_pending & " Days."

			  cash_listed = cash_listed + 1
			  y_pos = y_pos + 50

			  If pnd2_2nd_cash_prog_two <> "" Then
				  GroupBox 15, y_pos, 315, 45, "Cash Denial"
				  If pnd2_2nd_cash_status = "I" Then Text 20, y_pos + 15, 155, 10, "Cash Dened for NO INTERVIEW"
				  If pnd2_2nd_cash_status = "R" Then Text 20, y_pos + 15, 155, 10, "Cash Application WITHDRAWN"
				  Text 20, y_pos + 25, 155, 10, "Cash Program: " & pnd2_2nd_cash_prog_two
				  Text 185, y_pos + 15, 140, 10, "Cash application date: " & pnd2_2nd_appl_date
				  Text 185, y_pos + 25, 140, 10, "Cash has been pending for " & pnd2_2nd_days_pending & " Days."

				  cash_listed = cash_listed + 1
				  y_pos = y_pos + 50
			  End If
		  End If

		  If pnd2_snap_status = "I" or pnd2_snap_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "SNAP Denial"
			  If pnd2_snap_status = "I" Then Text 20, y_pos + 15, 155, 10, "SNAP Dened for NO INTERVIEW"
			  If pnd2_snap_status = "R" Then Text 20, y_pos + 15, 155, 10, "SNAP Application WITHDRAWN"
			  Text 185, y_pos + 15, 140, 10, "SNAP application date: " & pnd2_appl_date
			  Text 185, y_pos + 25, 140, 10, "SNAP has been pending for " & pnd2_days_pending & " Days."

			  y_pos = y_pos + 50
		  End If
		  If pnd2_2nd_snap_status = "I" or pnd2_2nd_snap_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "SNAP Denial"
			  If pnd2_2nd_snap_status = "I" Then Text 20, y_pos + 15, 155, 10, "SNAP Dened for NO INTERVIEW"
			  If pnd2_2nd_snap_status = "R" Then Text 20, y_pos + 15, 155, 10, "SNAP Application WITHDRAWN"
			  Text 185, y_pos + 15, 140, 10, "SNAP application date: " & pnd2_2nd_appl_date
			  Text 185, y_pos + 25, 140, 10, "SNAP has been pending for " & pnd2_2nd_days_pending & " Days."

			  y_pos = y_pos + 50
		  End If

		  If pnd2_emer_status = "I" or pnd2_emer_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "Emergency Denial"
			  If pnd2_emer_status = "I" Then Text 20, y_pos + 15, 155, 10, "EMER Dened for NO INTERVIEW"
			  If pnd2_emer_status = "R" Then Text 20, y_pos + 15, 155, 10, "EMER Application WITHDRAWN"
			  Text 185, y_pos + 15, 140, 10, "EMER application date: " & pnd2_appl_date
			  Text 185, y_pos + 25, 140, 10, "EMER has been pending for " & pnd2_days_pending & " Days."

			  y_pos = y_pos + 50
		  End If
		  If pnd2_2nd_emer_status = "I" or pnd2_2nd_emer_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "Emergency Denial"
			  If pnd2_2nd_emer_status = "I" Then Text 20, y_pos + 15, 155, 10, "EMER Dened for NO INTERVIEW"
			  If pnd2_2nd_emer_status = "R" Then Text 20, y_pos + 15, 155, 10, "EMER Application WITHDRAWN"
			  Text 185, y_pos + 15, 140, 10, "EMER application date: " & pnd2_2nd_appl_date
			  Text 185, y_pos + 25, 140, 10, "EMER has been pending for " & pnd2_2nd_days_pending & " Days."

			  y_pos = y_pos + 50
		  End If

		  If pnd2_grh_status = "I" or pnd2_grh_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "GRH Denial"
			  If pnd2_grh_status = "I" Then Text 20, y_pos + 15, 155, 10, "GRH Dened for NO INTERVIEW"
			  If pnd2_grh_status = "R" Then Text 20, y_pos + 15, 155, 10, "GRH Application WITHDRAWN"
			  Text 185, y_pos + 15, 140, 10, "GRH application date: " & pnd2_appl_date
			  Text 185, y_pos + 25, 140, 10, "GRH has been pending for " & pnd2_days_pending & " Days."

			  y_pos = y_pos + 50
		  End If
		  If pnd2_2nd_grh_status = "I" or pnd2_2nd_grh_status = "R" then
			  GroupBox 15, y_pos, 315, 45, "GRH Denial"
			  If pnd2_2nd_grh_status = "I" Then Text 20, y_pos + 15, 155, 10, "GRH Dened for NO INTERVIEW"
			  If pnd2_2nd_grh_status = "R" Then Text 20, y_pos + 15, 155, 10, "GRH Application WITHDRAWN"
			  Text 185, y_pos + 15, 140, 10, "GRH application date: " & pnd2_2nd_appl_date
			  Text 185, y_pos + 25, 140, 10, "GRH has been pending for " & pnd2_2nd_days_pending & " Days."

			  y_pos = y_pos + 50
		  End If
		  Text 20, y_pos+5, 135, 10, "Confirm that these denials are accurate:"
		  DropListBox 155, y_pos, 155, 45, "Indicate if the Denial is Accurate"+chr(9)+"Yes - denial is Accurate"+chr(9)+"No - I need to update the denial", denial_accurate
		  y_pos = y_pos + 20
		  ButtonGroup ButtonPressed
		    OkButton 230, y_pos, 50, 15
		    CancelButton 280, y_pos, 50, 15

		  ' y_pos = y_pos -
		  ' GroupBox 15, 25, 315, 45, "Cash Denial"
		  ' Text 20, 40, 155, 10, "Cash Dened for NO INTERVIEW"
		  ' Text 20, 50, 155, 10, "Cash Program: "
		  ' Text 185, 40, 140, 10, "Cash application date: "
		  ' Text 185, 50, 140, 10, "Cash has been pending for  Days."
		  ' GroupBox 15, 75, 315, 45, "Cash Denial"
		  ' Text 20, 90, 155, 10, "Cash Dened for NO INTERVIEW"
		  ' Text 20, 100, 155, 10, "Cash Program: "
		  ' Text 185, 90, 140, 10, "Cash application date: "
		  ' Text 185, 100, 140, 10, "Cash has been pending for  Days."
		  ' GroupBox 15, 125, 315, 45, "SNAP Denial"
		  ' Text 20, 140, 155, 10, "SNAP Application Withdrawn"
		  ' Text 185, 140, 140, 10, "SNAP application date: "
		  ' Text 185, 150, 140, 10, "SNAP has been pending for  Days."
		  ' GroupBox 15, 175, 315, 45, "EMER Denial"
		  ' Text 20, 190, 155, 10, "EMER Dened for NO INTERVIEW"
		  ' Text 20, 200, 155, 10, "EMER Program: "
		  ' Text 185, 190, 140, 10, "EMER application date: "
		  ' Text 185, 200, 140, 10, "EMER has been pending for  Days."
		  ' Text 20, 230, 135, 10, "Confirm that these denials are accurate:"
		  ' DropListBox 155, 225, 155, 45, "Indicate if the Denial is Accurate"+chr(9)+"Yes - denial is Accurate"+chr(9)+"No - I need to update the denial", denial_accurate
		  ' ButtonGroup ButtonPressed
		  '   OkButton 230, 245, 50, 15
		  '   CancelButton 280, 245, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If denial_accurate = "Indicate if the Denial is Accurate" Then MsgBox "*** Review the information on REPT/PND2 Denials ***" & vbCr & vbCr & "Ensure this is the intended result, this case will deny in the overnight process based on this information." & vbCr &vbCr & "Enter the answer for if the denials are accurate."

	Loop until denial_accurate <> "Indicate if the Denial is Accurate"

	If denial_accurate = "No - I need to update the denial" Then end_msg_info = end_msg_info & "CASE/NOTE has NOT been entered for REPT/PND2 denial as it was indicated the information was not accurate." & vbCr

	If denial_accurate = "Yes - denial is Accurate" Then

		If deny_app_one = True Then
			progs_denied_for_intv = ""
			progs_denied_for_wthdrw = ""

			appt_notc_date = ""
			nomi_date = ""

			If pnd2_cash_status = "I" Then
				If pnd2_cash_prog_one <> "" Then progs_denied_for_intv = progs_denied_for_intv & pnd2_cash_prog_one & ", "
				If pnd2_cash_prog_two <> "" Then progs_denied_for_intv = progs_denied_for_intv & pnd2_cash_prog_two & ", "
			End If
			If pnd2_snap_status = "I" Then progs_denied_for_intv = progs_denied_for_intv & "SNAP, "
			If pnd2_emer_status = "I" Then progs_denied_for_intv = progs_denied_for_intv & "EMER, "
			If pnd2_grh_status = "I" Then progs_denied_for_intv = progs_denied_for_intv & "GRH, "

			If pnd2_cash_status = "R" Then
				If pnd2_cash_prog_one <> "" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & pnd2_cash_prog_one & ", "
				If pnd2_cash_prog_two <> "" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & pnd2_cash_prog_two & ", "
			End If
			If pnd2_snap_status = "R" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & "SNAP, "
			If pnd2_emer_status = "R" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & "EMER, "
			If pnd2_grh_status = "R" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & "GRH, "

			If right(progs_denied_for_intv, 2) = ", " Then progs_denied_for_intv = left(progs_denied_for_intv, len(progs_denied_for_intv)-2)
			If right(progs_denied_for_wthdrw, 2) = ", " Then progs_denied_for_wthdrw = left(progs_denied_for_wthdrw, len(progs_denied_for_wthdrw)-2)

			If progs_denied_for_intv <> "" Then

				Call navigate_to_MAXIS_screen("CASE", "NOTE")       'First to case note to find what has ahppened'
				day_before_app = DateAdd("d", -1,pnd2_appl_date) 'will set the date one day prior to app date'

				note_row = 5            'resetting the variables on the loop
				note_date = ""
				note_title = ""
				appt_date = ""
				Do
					EMReadScreen note_date, 8, note_row, 6      'reading the note date
					EMReadScreen note_title, 55, note_row, 25   'reading the note header
					note_title = trim(note_title)
					IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then appt_notc_date = note_date
					IF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then appt_notc_date = note_date
					IF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then appt_notc_date = note_date

					IF note_title = "~ Client missed application interview, NOMI sent via sc" then nomi_date = note_date
					IF left(note_title, 32) = "**Client missed SNAP interview**" then nomi_date = note_date
					IF left(note_title, 32) = "**Client missed CASH interview**" then nomi_date = note_date
					IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then nomi_date = note_date
					IF note_title = "~ Client has not completed application interview, NOMI" then nomi_date = note_date
					IF note_title = "~ Client has not completed CASH APP interview, NOMI sen" then nomi_date = note_date
					IF note_title = "* A notice was previously sent to client with detail ab" then nomi_date = note_date

					IF note_date = "        " then Exit Do
					note_row = note_row + 1
					IF note_row = 19 THEN
						PF8
						note_row = 5
					END IF
					EMReadScreen next_note_date, 8, note_row, 6
					IF next_note_date = "        " then Exit Do
				Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
				PF3

				end_msg_info = end_msg_info & "CASE/NOTE entered for Denial of " & pnd2_appl_date & " application of " & progs_denied_for_intv & " - denied on REPT/PND2 for no Interview." & vbCr
				Call start_a_blank_CASE_NOTE

				Call write_variable_in_CASE_NOTE("DENIAL of " & pnd2_appl_date & " Application for No Interview: " & progs_denied_for_intv)
				Call write_variable_in_CASE_NOTE("REPT/PND2 has been updated to deny this case in an overnight system process.")
				Call write_bullet_and_variable_in_CASE_NOTE("Application Date", pnd2_appl_date)
				Call write_bullet_and_variable_in_CASE_NOTE("Programs to Deny", progs_denied_for_intv)
				Call write_bullet_and_variable_in_CASE_NOTE("Notice of Interview Sent Date", appt_notc_date)
				Call write_bullet_and_variable_in_CASE_NOTE("NOMI (Notice of Missed Interview) Send Date", nomi_date)
				Call write_variable_in_CASE_NOTE("---")
				Call write_variable_in_CASE_NOTE(worker_signature)

			End If

			If progs_denied_for_wthdrw <> "" Then
				end_msg_info = end_msg_info & "CASE/NOTE entered for Denial of " & pnd2_appl_date & " application of " & progs_denied_for_wthdrw & " - denied on REPT/PND2 for Withdraw of Request." & vbCr

				Call start_a_blank_CASE_NOTE

				Call write_variable_in_CASE_NOTE("DENIAL of " & pnd2_appl_date & " Application by resident Request: " & progs_denied_for_wthdrw)
				Call write_variable_in_CASE_NOTE("REPT/PND2 has been updated to deny this case in an overnight system process.")
				Call write_bullet_and_variable_in_CASE_NOTE("Application Date", pnd2_appl_date)
				Call write_bullet_and_variable_in_CASE_NOTE("Programs to Deny", progs_denied_for_intv)
				Call write_variable_in_CASE_NOTE("* Resident requested to withdraw application.")
				Call write_variable_in_CASE_NOTE("---")
				Call write_variable_in_CASE_NOTE(worker_signature)
			End If

		End If

		If deny_app_two = True Then
			progs_denied_for_intv = ""
			progs_denied_for_wthdrw = ""

			appt_notc_date = ""
			nomi_date  ""

			If pnd2_2nd_cash_status = "I" Then
				If pnd2_2nd_cash_prog_one <> "" Then progs_denied_for_intv = progs_denied_for_intv & pnd2_2nd_cash_prog_one & ", "
				If pnd2_2nd_cash_prog_two <> "" Then progs_denied_for_intv = progs_denied_for_intv & pnd2_2nd_cash_prog_two & ", "
			End If
			If pnd2_2nd_snap_status = "I" Then progs_denied_for_intv = progs_denied_for_intv & "SNAP, "
			If pnd2_2nd_emer_status = "I" Then progs_denied_for_intv = progs_denied_for_intv & "EMER, "
			If pnd2_2nd_grh_status = "I" Then progs_denied_for_intv = progs_denied_for_intv & "GRH, "

			If pnd2_2nd_cash_status = "R" Then
				If pnd2_2nd_cash_prog_one <> "" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & pnd2_2nd_cash_prog_one & ", "
				If pnd2_2nd_cash_prog_two <> "" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & pnd2_2nd_cash_prog_two & ", "
			End If
			If pnd2_2nd_snap_status = "R" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & "SNAP, "
			If pnd2_2nd_emer_status = "R" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & "EMER, "
			If pnd2_2nd_grh_status = "R" Then progs_denied_for_wthdrw = progs_denied_for_wthdrw & "GRH, "

			If right(progs_denied_for_intv, 2) = ", " Then progs_denied_for_intv = left(progs_denied_for_intv, len(progs_denied_for_intv)-2)
			If right(progs_denied_for_wthdrw, 2) = ", " Then progs_denied_for_wthdrw = left(progs_denied_for_wthdrw, len(progs_denied_for_wthdrw)-2)

			If progs_denied_for_intv <> "" Then

				Call navigate_to_MAXIS_screen("CASE", "NOTE")       'First to case note to find what has ahppened'
				day_before_app = DateAdd("d", -1,pnd2_2nd_appl_date) 'will set the date one day prior to app date'

				note_row = 5            'resetting the variables on the loop
				note_date = ""
				note_title = ""
				appt_date = ""
				Do
					EMReadScreen note_date, 8, note_row, 6      'reading the note date
					EMReadScreen note_title, 55, note_row, 25   'reading the note header
					note_title = trim(note_title)
					IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then appt_notc_date = note_date
					IF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then appt_notc_date = note_date
					IF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then appt_notc_date = note_date

					IF note_title = "~ Client missed application interview, NOMI sent via sc" then nomi_date = note_date
					IF left(note_title, 32) = "**Client missed SNAP interview**" then nomi_date = note_date
					IF left(note_title, 32) = "**Client missed CASH interview**" then nomi_date = note_date
					IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then nomi_date = note_date
					IF note_title = "~ Client has not completed application interview, NOMI" then nomi_date = note_date
					IF note_title = "~ Client has not completed CASH APP interview, NOMI sen" then nomi_date = note_date
					IF note_title = "* A notice was previously sent to client with detail ab" then nomi_date = note_date

					IF note_date = "        " then Exit Do
					note_row = note_row + 1
					IF note_row = 19 THEN
						PF8
						note_row = 5
					END IF
					EMReadScreen next_note_date, 8, note_row, 6
					IF next_note_date = "        " then Exit Do
				Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
				PF3

				end_msg_info = end_msg_info & "CASE/NOTE entered for Denial of " & pnd2_2nd_appl_date & " application of " & progs_denied_for_intv & " - denied on REPT/PND2 for no Interview." & vbCr
				Call start_a_blank_CASE_NOTE

				Call write_variable_in_CASE_NOTE("DENIAL of " & pnd2_2nd_appl_date & " Application for No Interview: " & progs_denied_for_intv)
				Call write_variable_in_CASE_NOTE("REPT/PND2 has been updated to deny this case in an overnight system process.")
				Call write_bullet_and_variable_in_CASE_NOTE("Application Date", pnd2_2nd_appl_date)
				Call write_bullet_and_variable_in_CASE_NOTE("Programs to Deny", progs_denied_for_intv)
				Call write_bullet_and_variable_in_CASE_NOTE("Notice of Interview Sent Date", appt_notc_date)
				Call write_bullet_and_variable_in_CASE_NOTE("NOMI (Notice of Missed Interview) Send Date", nomi_date)
				Call write_variable_in_CASE_NOTE("---")
				Call write_variable_in_CASE_NOTE(worker_signature)

			End If

			If progs_denied_for_wthdrw <> "" Then
				end_msg_info = end_msg_info & "CASE/NOTE entered for Denial of " & pnd2_2nd_appl_date & " application of " & progs_denied_for_wthdrw & " - denied on REPT/PND2 for Withdraw of Request." & vbCr
				Call start_a_blank_CASE_NOTE

				Call write_variable_in_CASE_NOTE("DENIAL of " & pnd2_2nd_appl_date & " Application by resident Request: " & progs_denied_for_wthdrw)
				Call write_variable_in_CASE_NOTE("REPT/PND2 has been updated to deny this case in an overnight system process.")
				Call write_bullet_and_variable_in_CASE_NOTE("Application Date", pnd2_2nd_appl_date)
				Call write_bullet_and_variable_in_CASE_NOTE("Programs to Deny", progs_denied_for_intv)
				Call write_variable_in_CASE_NOTE("* Resident requested to withdraw application.")
				Call write_variable_in_CASE_NOTE("---")
				Call write_variable_in_CASE_NOTE(worker_signature)
			End If
		End If

	End If


End If


' "- 04/22 . . . Entitlement:    $ "250
' "              Prorated:       $ "150 (Prorated from 04/14/2022)
' "              Recoupment:   - $ " 25
' "              Issued to Resident:   $ "125

'TODO - figure out how to read for the possible errors.'

' For approval_month = 0 to UBound(SNAP_ELIG_APPROVALS)
' 	For snap_memb = 0 to UBound(SNAP_ELIG_APPROVALS(approval_month).snap_elig_ref_numbs)
' 		MsgBox SNAP_ELIG_APPROVALS(approval_month).elig_footer_month & "/" & SNAP_ELIG_APPROVALS(approval_month).elig_footer_year & vbCr & SNAP_ELIG_APPROVALS(approval_month).snap_elig_ref_numbs(snap_memb) & vbCr & SNAP_ELIG_APPROVALS(approval_month).snap_elig_membs_eligibility(snap_memb)
' 	Next
' Next

If pnd2_display_limit_hit = True AND denials_found_on_pnd2 = False Then end_msg_info = end_msg_info & vbCr & vbCr & "The script could not read REPT/PND2 because the X-Number it is in has hit the MAXIS REPT/PND2 display limit. If you are trying to deny the case via REPT/PND2, the case will need to be in an X-Number that is not at the REPT/PND2 display limit."


Call script_end_procedure_with_error_report("All approval information has been reviewed." & vbCr & vbCr & end_msg_info)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------
