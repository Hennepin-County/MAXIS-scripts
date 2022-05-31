'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
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
' call run_from_GitHub(script_repository & "application-received.vbs")

const assign_worker_col         = 1
const assign_case_numb_col      = 2
const assign_case_name_col      = 3
const assign_cash_col           = 4
const assign_snap_col           = 5
const assign_op_col             = 6
const assign_no_revw_form_col   = 7
const assign_supp_col           = 8
const assign_case_correct_col   = 9
const assign_notes_comments_col = 10
const assign_tier_two_notes_col = 11
const assign_tier_two_worker_col= 12
const assign_tier_two_process_col = 13


'COLUMNS'
const det_case_numb_col 		=  1
const det_process_col 			=  2
const det_issued_fs_f_col 		=  3
const det_issued_fs_s_col 		=  4
const det_issued_mf_mf_col 		=  5
const det_issued_mf_fs_f_col 	=  6
const det_issued_mf_fs_s_col 	=  7
const det_issued_mf_hg_col 		=  8
const det_form_col 				=  9
const det_form_date_col 		= 10
const det_intv_col 				= 11
const det_intv_date_col 		= 12
const det_verifs_col 			= 13
const det_process_complete_col 	= 14
const det_op_fs_f_col 			= 15
const det_op_fs_s_col 			= 16
const det_op_mf_mf_col 			= 17
const det_op_mf_fs_f_col 		= 18
const det_op_mf_fs_s_col 		= 19
const det_op_mf_hg_col 			= 20
const det_supp_fs_f_col 		= 21
const det_supp_fs_s_col 		= 22
const det_supp_mf_mf_col 		= 23
const det_supp_mf_fs_f_col 		= 24
const det_supp_mf_fs_s_col 		= 25
const det_supp_mf_hg_col 		= 26
const det_orig_earned_income_col 		= 27
const det_orig_unearned_income_col 		= 28
const det_orig_total_income_col 		= 29
const det_orig_total_ded_col 			= 30
const det_orig_net_income_col 			= 31
const det_orig_housing_cost_col 		= 32
const det_orig_utility_cost_col 		= 33
const det_orig_total_shel_cost_col 		= 34
const det_orig_net_adj_income_col 		= 35
const det_orig_hh_size_col 				= 36
const det_orig_snap_benefit_col 		= 37
const det_correct_earned_income_col 	= 38
const det_correct_unearned_income_col 	= 39
const det_correct_total_income_col 		= 40
const det_correct_total_ded_col 		= 41
const det_correct_net_income_col 		= 42
const det_correct_housing_cost_col 		= 43
const det_correct_utility_cost_col 		= 44
const det_correct_total_shel_cost_col 	= 45
const det_correct_net_adj_income_col 	= 46
const det_correct_hh_size_col 			= 47
const det_snap_proration_col			= 48
const det_correct_snap_benefit_col 		= 49

const det_orig_mf_caregivers_col 			= 50 'Orig Caregivers
const det_orig_mf_children_col 				= 51 'Orig Children
const det_orig_mf_earned_income_col 		= 52 'Orig MF Earned Income
const det_orig_mf_ei_deductions_col 		= 53 'Orig EI Disregards/Deductions
const det_orig_mf_net_ei_col 				= 54 'Orig Net Earned Income
const det_orig_mf_family_wage_level_col 	= 55 'Orig Family Wage Level
const det_orig_mf_difference_col 			= 56 'Orig Difference
const det_orig_mf_transitional_stndrd_col 	= 57 'Orig Trasitional Standard
const det_orig_mf_diff_or_trns_stndrd_col 	= 58 'Difference or Transitional Standard
const det_orig_mf_unearned_income_col 		= 59 'Orig MF Unearned Income
const det_orig_mf_unea_deductions_col 		= 60 'Orig Ded/Disrgd
const det_orig_mf_deemed_income_col 		= 61 'Orig Deemed Income
const det_orig_mf_cs_exclusion_col 			= 62 'Orig CS Exclusion
const det_orig_mf_subsidy_col 				= 63 'Orig Subsidy
const det_orig_mf_unmet_need_col 			= 64 'Orig Unmet Need
const det_orig_mf_mf_amt_col 				= 65 'Orig MF-MF
const det_orig_mf_fs_amt_col 				= 66 'Orig MF-FS
const det_orig_mf_hg_amt_col 				= 67 'Orig MF-HG
const det_correct_mf_caregivers_col 		= 68 'Correct Caregivers
const det_correct_mf_children_col 			= 69 'Correct Children
const det_correct_mf_earned_income_col 		= 70 'Correct MF Earned Income
const det_correct_mf_ei_deductions_col 		= 71 'Correct EI Disregards/Deductions
const det_correct_mf_net_ei_col 			= 72 'Correct Net Earned Income
const det_correct_mf_family_wage_level_col 	= 73 'Correct Family Wage Level
const det_correct_mf_difference_col 		= 74 'Correct Difference
const det_correct_mf_transitional_stndrd_col = 75 'Correct Trasitional Standard
const det_correct_mf_diff_or_trns_stndrd_col = 76 'Correct or Transitional Standard
const det_correct_mf_unearned_income_col 	= 77 'Correct MF Unearned Income
const det_correct_mf_unea_deductions_col 	= 78 'Correct Ded/Disrgd
const det_correct_mf_deemed_income_col 		= 79 'Correct Deemed Income
const det_correct_mf_cs_exclusion_col 		= 80 'Correct CS Exclusion
const det_correct_mf_subsidy_col 			= 81 'Correct Subsidy
const det_mf_proration_date_col 			= 82 'MFIP Proration Date
const det_correct_mf_unmet_need_col 		= 83 'Correct Unmet Need
const det_correct_mf_mf_amt_col 			= 84 'Correct MF-MF
const det_correct_mf_fs_amt_col 			= 85 'Correct MF-FS
const det_correct_mf_hg_amt_col 			= 86 'Correct MF-HG


const det_snap_pdf_link_col 				= 87
const det_mfip_pdf_link_col 				= 88


const rept_case_numb_col 		=  14
const rept_process_col 			=  15
const rept_issued_fs_f_col 		=  16
const rept_issued_fs_s_col 		=  17
const rept_issued_mf_mf_col 		=  18
const rept_issued_mf_fs_f_col 	=  19
const rept_issued_mf_fs_s_col 	=  20
const rept_issued_mf_hg_col 		=  21
const rept_form_col 				=  22
const rept_form_date_col 		= 23
const rept_intv_col 				= 24
const rept_intv_date_col 		= 25
const rept_verifs_col 			= 26
const rept_process_complete_col 	= 27
const rept_op_fs_f_col 			= 28
const rept_op_fs_s_col 			= 29
const rept_op_mf_mf_col 			= 30
const rept_op_mf_fs_f_col 		= 31
const rept_op_mf_fs_s_col 		= 32
const rept_op_mf_hg_col 			= 33
const rept_supp_fs_f_col 		= 34
const rept_supp_fs_s_col 		= 35
const rept_supp_mf_mf_col 		= 36
const rept_supp_mf_fs_f_col 		= 37
const rept_supp_mf_fs_s_col 		= 38
const rept_supp_mf_hg_col 		= 39
const rept_orig_earned_income_col 		= 40
const rept_orig_unearned_income_col 		= 41
const rept_orig_total_income_col 		= 42
const rept_orig_total_ded_col 			= 43
const rept_orig_net_income_col 			= 44
const rept_orig_housing_cost_col 		= 45
const rept_orig_utility_cost_col 		= 46
const rept_orig_total_shel_cost_col 		= 47
const rept_orig_net_adj_income_col 		= 48
const rept_orig_hh_size_col 				= 49
const rept_orig_snap_benefit_col 		= 50
const rept_correct_earned_income_col 	= 51
const rept_correct_unearned_income_col 	= 52
const rept_correct_total_income_col 		= 53
const rept_correct_total_ded_col 		= 54
const rept_correct_net_income_col 		= 55
const rept_correct_housing_cost_col 		= 56
const rept_correct_utility_cost_col 		= 57
const rept_correct_total_shel_cost_col 	= 58
const rept_correct_net_adj_income_col 	= 59
const rept_correct_hh_size_col 			= 60
const rept_snap_proration_col			= 61
const rept_correct_snap_benefit_col 		= 62

const rept_orig_mf_caregivers_col 			= 63
const rept_orig_mf_children_col 				= 64
const rept_orig_mf_earned_income_col 		= 65
const rept_orig_mf_ei_deductions_col 		= 66
const rept_orig_mf_net_ei_col 				= 67
const rept_orig_mf_family_wage_level_col 	= 68
const rept_orig_mf_difference_col 			= 69
const rept_orig_mf_transitional_stndrd_col 	= 70
const rept_orig_mf_diff_or_trns_stndrd_col 	= 71
const rept_orig_mf_unearned_income_col 		= 72
const rept_orig_mf_unea_deductions_col 		= 73
const rept_orig_mf_deemed_income_col 		= 74
const rept_orig_mf_cs_exclusion_col 			= 75
const rept_orig_mf_subsidy_col 				= 76
const rept_orig_mf_unmet_need_col 			= 77
const rept_orig_mf_mf_amt_col 				= 78
const rept_orig_mf_fs_amt_col 				= 79
const rept_orig_mf_hg_amt_col 				= 80
const rept_correct_mf_caregivers_col 		= 81
const rept_correct_mf_children_col 			= 82
const rept_correct_mf_earned_income_col 		= 83
const rept_correct_mf_ei_deductions_col 		= 84
const rept_correct_mf_net_ei_col 			= 85
const rept_correct_mf_family_wage_level_col 	= 86
const rept_correct_mf_difference_col 		= 87
const rept_correct_mf_transitional_stndrd_col = 88
const rept_correct_mf_diff_or_trns_stndrd_col = 89
const rept_correct_mf_unearned_income_col 	= 90
const rept_correct_mf_unea_deductions_col 	= 91
const rept_correct_mf_deemed_income_col 		= 92
const rept_correct_mf_cs_exclusion_col 		= 93
const rept_correct_mf_subsidy_col 			= 94
const rept_mf_proration_date_col 			= 95
const rept_correct_mf_unmet_need_col 		= 96
const rept_correct_mf_mf_amt_col 			= 97
const rept_correct_mf_fs_amt_col 			= 98
const rept_correct_mf_hg_amt_col 			= 99


const rept_snap_pdf_link_col 				= 87
const rept_mfip_pdf_link_col 				= 88






const case_number_const                                     = 00
const feb_process_const                                     = 01
const fed_benefit_amt_const                                 = 02
const state_benefit_amt_const                               = 03
const mfip_MF_MF_issued_amt_const                           = 04
const mfip_MF_FS_F_issued_amt_const                         = 05
const mfip_MF_FS_S_issued_amt_const                         = 06
const mfip_MF_HG_issued_amt_const                           = 07
const form_received_const                                   = 08
const form_received_date_const                              = 09
const interview_information_const                           = 10
const interview_date_const                                  = 11
const verifs_received_const                                 = 12
const process_complete_const                                = 13
const SNAP_fed_op_const                                     = 14
const SNAP_state_op_const                                   = 15
const mfip_cash_overpayment_amt_const                       = 16
const mfip_food_f_overpayment_const                         = 17
const mfip_food_s_overpayment_const                         = 18
const mfip_hg_overpayment_amt_const                         = 19
const SNAP_fed_supp_const                                   = 20
const SNAP_state_supp_const                                 = 21
const mfip_cash_supplement_amt_const                        = 22
const mfip_food_s_supplement_const                          = 23
const mfip_food_f_supplement_const                          = 24
const mfip_hg_supplement_amt_const                          = 25
const earned_income_budgeted_amt_const                      = 26
const unearned_budgeted_amt_const                           = 27
const total_income_budgeted_amt_const                       = 28
const total_deduction_budgeted_amt_const                    = 29
const net_income_budgeted_amt_const                         = 30
const total_housing_cost_budgeted_amt_const                 = 31
const utilities_budgeted_amt_const                          = 32
const total_shelter_cost_budgeted_amt_const                 = 33
const net_adj_income_budgeted_amt_const                     = 34
const budgeted_hh_size_const                                = 35
const snap_issued_amt_const                                 = 36
const earned_income_correct_amt_const                       = 37
const unearned_correct_amt_const                            = 38
const total_income_correct_amt_const                        = 39
const total_deduction_correct_amt_const                     = 40
const net_income_correct_amt_const                          = 41
const total_housing_cost_correct_amt_const                  = 42
const utilities_correct_amt_const                           = 43
const total_shelter_cost_correct_amt_const                  = 44
const net_adj_income_correct_amt_const                      = 45
const correct_hh_size_const                                 = 46
const snap_proration_date_const                             = 47
const snap_correct_amt_const                                = 48
const mfip_budgeted_caregivers_const                        = 49
const mfip_budgeted_children_const 		                    = 50
const mfip_orig_gross_total_earned_income_const             = 51
const mfip_orig_deductions_earned_const 		            = 52
const mfip_budgeted_earned_income_const 			        = 53
const mf_orig_fwl_const                                     = 54
const mf_orig_fwl_diff_const  			                    = 55
const mf_orig_ts_const 			  			                = 56
const mf_orig_diff_or_ts_const 		  			            = 57
const mfip_orig_deductions_unearned_const  			        = 58
const deemed_income_budgeted_amt_const 		  			    = 59
const mfip_budgeted_unearned_income_const  			        = 60
const cses_exclusion_budgeted_amt_const  			        = 61
const mfip_subsidy_tribal_amt_const   			            = 62
const mfip_total_issued_amt_const 	  			            = 63
' const mfip_MF_MF_issued_amt_const 		  			        = 64
const mfip_MF_FS_issued_amt_const 	                        = 65
' const mfip_MF_HG_issued_amt_const 	                        = 66
const correct_caregiver_const 		                        = 67
const correct_children_const 		                        = 68
const total_correct_mfip_earned_income_const                = 69
const total_correct_mfip_earned_deductions_and_disreagards_const    = 70
const total_correct_mfip_net_earned_income_const            = 71
const familY_wage_level_const                               = 72
const wage_level_difference_const                           = 73
const full_mfip_standard_const 		                        = 74
const difference_or_transitional_const                      = 75
const total_correct_mfip_unearned_income_const              = 76
const total_correct_mfip_unearned_deductions_and_disreagards_const  = 77
const correct_mfip_deemed_amt_const                         = 78
const correct_mfip_cses_exclusion_const                     = 79
const mfip_proration_date_const 	                        = 80
const prorated_unmet_need_const 		                    = 81
const mfip_correct_cash_portion_const 		                = 82
const mfip_correct_food_portion_const 			            = 83
const mfip_correct_hg_portion_const 				        = 84
const snap_pdf_excel_cell_info_const                        = 85
const mfip_pdf_excel_cell_info_const                        = 86
const tier_two_detail_info_last_const                       = 87


'END FUNCTIONS LIBRARY BLOCK================================================================================================
EMConnect""

' const case_numb = 0
' const app_date  = 1
' const found     = 2
' const last_const = 3

Dim TIER_TWO_REVIEW_DETAIL_ARRAY()
ReDim TIER_TWO_REVIEW_DETAIL_ARRAY(tier_two_detail_info_last_const, 0)

Dim SQL_LIST_ARRAY()
ReDim SQL_LIST_ARRAY(last_const, 0)

bobi_cases_string = "~"
sql_cases_string = "~"

array_counter = 0

' Open the days excel
excel_details_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Tracking Details - Yeng.xlsx"
Call excel_open(excel_details_file_path, False, False, ObjDetailsExcel, objDetailWorkbook)

total_excel_row = 3
Do
    ReDim Preserve TIER_TWO_REVIEW_DETAIL_ARRAY(tier_two_detail_info_last_const, array_counter)


    TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_case_numb_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(feb_process_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_process_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(fed_benefit_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(state_benefit_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_mf_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_F_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_S_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_hg_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_form_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_form_date_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(interview_information_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_intv_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(interview_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_intv_date_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(verifs_received_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_verifs_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(process_complete_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_process_complete_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_op_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_op_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_overpayment_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_mf_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_overpayment_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_overpayment_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_overpayment_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_hg_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_supp_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_supp_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_supplement_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_mf_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_supplement_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_supplement_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_supplement_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_hg_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_total_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_total_ded_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_net_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_housing_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_utility_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_total_shel_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_net_adj_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(budgeted_hh_size_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_hh_size_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_snap_benefit_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_total_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_total_ded_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_net_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_housing_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_utility_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_total_shel_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_net_adj_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_hh_size_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_hh_size_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_proration_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_snap_proration_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_snap_benefit_col).Value)

    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_caregivers_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_caregivers_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_children_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_children_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_gross_total_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_earned_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_ei_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_net_ei_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_family_wage_level_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_diff_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_difference_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_ts_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_transitional_stndrd_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_diff_or_ts_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_diff_or_trns_stndrd_col).Value)

    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_unearned_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_unea_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(deemed_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_deemed_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_unearned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(cses_exclusion_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_cs_exclusion_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_subsidy_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_total_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_unmet_need_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_mf_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_fs_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_hg_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_caregiver_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_caregivers_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_children_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_children_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_deductions_and_disreagards_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_ei_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_net_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_net_ei_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(familY_wage_level_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_family_wage_level_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(wage_level_difference_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_difference_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(full_mfip_standard_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_transitional_stndrd_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(difference_or_transitional_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_diff_or_trns_stndrd_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_deductions_and_disreagards_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_unea_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_deemed_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_deemed_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_cses_exclusion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_cs_exclusion_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_subsidy_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_proration_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_mf_proration_date_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(prorated_unmet_need_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_unmet_need_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_cash_portion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_mf_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_food_portion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_fs_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_hg_portion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_hg_amt_col).Value)

    ' If snap_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_snap_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & snap_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_snap_file_name & chr(34) & ")"
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_pdf_excel_cell_info_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_snap_pdf_link_col).Value)
    ' If mfip_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_mfip_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & mfip_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_mfip_file_name & chr(34) & ")"
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_pdf_excel_cell_info_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_mfip_pdf_link_col).Value)


    array_counter = array_counter + 1
    total_excel_row = total_excel_row + 1
    next_case_numb = trim(ObjDetailsExcel.Cells(total_excel_row, det_case_numb_col).Value)
Loop until next_case_numb = ""

ObjDetailsExcel.ActiveWorkbook.Close

ObjDetailsExcel.Application.Quit
ObjDetailsExcel.Quit

excel_details_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\AutoClose Pause Tier Two Tracking Details - Mandora.xlsx"
Call excel_open(excel_details_file_path, False, False, ObjDetailsExcel, objDetailWorkbook)

total_excel_row = 3
Do
    ReDim Preserve TIER_TWO_REVIEW_DETAIL_ARRAY(tier_two_detail_info_last_const, array_counter)


    TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_case_numb_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(feb_process_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_process_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(fed_benefit_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(state_benefit_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_mf_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_F_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_S_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_issued_mf_hg_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_form_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_form_date_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(interview_information_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_intv_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(interview_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_intv_date_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(verifs_received_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_verifs_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(process_complete_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_process_complete_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_op_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_op_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_overpayment_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_mf_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_overpayment_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_overpayment_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_overpayment_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_op_mf_hg_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_supp_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_supp_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_supplement_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_mf_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_supplement_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_fs_s_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_supplement_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_fs_f_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_supplement_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_supp_mf_hg_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_total_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_total_ded_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_net_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_housing_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_utility_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_total_shel_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_net_adj_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(budgeted_hh_size_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_hh_size_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_snap_benefit_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_total_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_total_ded_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_net_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_housing_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_utility_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_total_shel_cost_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_net_adj_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_hh_size_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_hh_size_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_proration_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_snap_proration_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_correct_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_snap_benefit_col).Value)

    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_caregivers_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_caregivers_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_children_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_children_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_gross_total_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_earned_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_ei_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_net_ei_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_family_wage_level_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_diff_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_difference_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_ts_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_transitional_stndrd_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_diff_or_ts_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_diff_or_trns_stndrd_col).Value)

    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_unearned_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_unea_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(deemed_income_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_deemed_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_unearned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(cses_exclusion_budgeted_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_cs_exclusion_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_subsidy_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_total_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_unmet_need_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_mf_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_fs_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_orig_mf_hg_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_caregiver_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_caregivers_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_children_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_children_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_earned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_deductions_and_disreagards_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_ei_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_net_earned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_net_ei_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(familY_wage_level_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_family_wage_level_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(wage_level_difference_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_difference_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(full_mfip_standard_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_transitional_stndrd_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(difference_or_transitional_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_diff_or_trns_stndrd_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_income_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_unearned_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_deductions_and_disreagards_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_unea_deductions_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_deemed_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_deemed_income_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_cses_exclusion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_cs_exclusion_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_subsidy_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_proration_date_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_mf_proration_date_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(prorated_unmet_need_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_unmet_need_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_cash_portion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_mf_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_food_portion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_fs_amt_col).Value)
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_hg_portion_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_correct_mf_hg_amt_col).Value)

    ' If snap_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_snap_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & snap_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_snap_file_name & chr(34) & ")"
    TIER_TWO_REVIEW_DETAIL_ARRAY(snap_pdf_excel_cell_info_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_snap_pdf_link_col).Value)
    ' If mfip_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_mfip_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & mfip_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_mfip_file_name & chr(34) & ")"
    TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_pdf_excel_cell_info_const, array_counter) = trim(ObjDetailsExcel.Cells(total_excel_row, det_mfip_pdf_link_col).Value)


    array_counter = array_counter + 1
    total_excel_row = total_excel_row + 1
    next_case_numb = trim(ObjDetailsExcel.Cells(total_excel_row, det_case_numb_col).Value)
Loop until next_case_numb = ""

ObjDetailsExcel.ActiveWorkbook.Close

ObjDetailsExcel.Application.Quit
ObjDetailsExcel.Quit

' MsgBox "array_counter - " & array_counter

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "Case information"
ObjExcel.ActiveSheet.Name = "Case information"

'Setting the first 4 col as worker, case number, name, and APPL date

ObjExcel.Cells(2, assign_worker_col).Value = "WORKER"
ObjExcel.Cells(2, assign_case_numb_col).Value = "CASE NUMBER"
ObjExcel.Cells(2, assign_case_name_col).Value = "NAME"
ObjExcel.Cells(2, assign_cash_col).Value = "CASH?"
ObjExcel.Cells(2, assign_snap_col).Value = "SNAP?"
ObjExcel.Cells(2, assign_op_col).Value = "Overpayment amount"
ObjExcel.Cells(2, assign_no_revw_form_col).Value = "No review form received"
ObjExcel.Cells(2, assign_supp_col).Value = "Underpayment Amount"
ObjExcel.Cells(2, assign_case_correct_col).Value = "No change in budget case,  closed correctly, benefits issued correctly, etc."
ObjExcel.Cells(2, assign_notes_comments_col).Value = "Notes/comments-Please list if a person is coded as a 6 month but they are not a 6 month report ect…."
ObjExcel.Cells(2, assign_tier_two_notes_col).Value = "Tier Two Notes/Tracking"
ObjExcel.Cells(2, assign_tier_two_worker_col).Value = "Tier Two Worker"
ObjExcel.Cells(2, assign_tier_two_process_col).Value = "Tier Two Process"

ObjExcel.Cells(2, rept_case_numb_col).Value = "Case Number"
ObjExcel.Cells(2, rept_process_col).Value = "Process"
ObjExcel.Cells(2, rept_issued_fs_f_col).Value = "FS - F - Issued"
ObjExcel.Cells(2, rept_issued_fs_s_col).Value = "FS - S - Issued"
ObjExcel.Cells(2, rept_issued_mf_mf_col).Value = "MF - MF - Issued"
ObjExcel.Cells(2, rept_issued_mf_fs_f_col).Value = "MF - FS - F - Issued"
ObjExcel.Cells(2, rept_issued_mf_fs_s_col).Value = "MF - FS - S - Issued"
ObjExcel.Cells(2, rept_issued_mf_hg_col).Value = "MF - HG - Issued"
ObjExcel.Cells(2, rept_form_col).Value = "Form Received"
ObjExcel.Cells(2, rept_form_date_col).Value = "Date Form Received"
ObjExcel.Cells(2, rept_intv_col).Value = "Interview"
ObjExcel.Cells(2, rept_intv_date_col).Value = "Interview Date"
ObjExcel.Cells(2, rept_verifs_col).Value = "Verifications Received"
ObjExcel.Cells(2, rept_process_complete_col).Value = "Was the Process Complete"
ObjExcel.Cells(2, rept_op_fs_f_col).Value = "FS - F - OP"
ObjExcel.Cells(2, rept_op_fs_s_col).Value = "FS - S - OP"
ObjExcel.Cells(2, rept_op_mf_mf_col).Value = "MF - MF - OP"
ObjExcel.Cells(2, rept_op_mf_fs_f_col).Value = "MF - FS - F - OP"
ObjExcel.Cells(2, rept_op_mf_fs_s_col).Value = "MF - FS - S - OP"
ObjExcel.Cells(2, rept_op_mf_hg_col).Value = "MF - HG - OP"
ObjExcel.Cells(2, rept_supp_fs_f_col).Value = "FS - F - Supp"
ObjExcel.Cells(2, rept_supp_fs_s_col).Value = "FS - S - Supp"
ObjExcel.Cells(2, rept_supp_mf_mf_col).Value = "MF - MF - Supp"
ObjExcel.Cells(2, rept_supp_mf_fs_f_col).Value = "MF - FS - F - Supp"
ObjExcel.Cells(2, rept_supp_mf_fs_s_col).Value = "MF - FS - S - Supp"
ObjExcel.Cells(2, rept_supp_mf_hg_col).Value = "MF - HG - Supp"
ObjExcel.Cells(2, rept_orig_earned_income_col).Value = "Orig Earned Income"
ObjExcel.Cells(2, rept_orig_unearned_income_col).Value = "Orig Unearned Income"
ObjExcel.Cells(2, rept_orig_total_income_col).Value = "Orig Total Income"
ObjExcel.Cells(2, rept_orig_total_ded_col).Value = "Orig Total Deductions"
ObjExcel.Cells(2, rept_orig_net_income_col).Value = "Orig Net Income"
ObjExcel.Cells(2, rept_orig_housing_cost_col).Value = "Orig Housing Expense"
ObjExcel.Cells(2, rept_orig_utility_cost_col).Value = "Orig Utilty Expense"
ObjExcel.Cells(2, rept_orig_total_shel_cost_col).Value = "Orig Total Shelter Costs"
ObjExcel.Cells(2, rept_orig_net_adj_income_col).Value = "Orig Net Adj Income"
ObjExcel.Cells(2, rept_orig_hh_size_col).Value = "Orig HH Size"
ObjExcel.Cells(2, rept_orig_snap_benefit_col).Value = "Orig SNAP Benefit"
ObjExcel.Cells(2, rept_correct_earned_income_col).Value = "Correct Earned Income"
ObjExcel.Cells(2, rept_correct_unearned_income_col).Value = "Correct Unearned Income"
ObjExcel.Cells(2, rept_correct_total_income_col).Value = "Correct Total Income"
ObjExcel.Cells(2, rept_correct_total_ded_col).Value = "Correct Total Deductions"
ObjExcel.Cells(2, rept_correct_net_income_col).Value = "Correct Net Income"
ObjExcel.Cells(2, rept_correct_housing_cost_col).Value = "Correct Housing Expense"
ObjExcel.Cells(2, rept_correct_utility_cost_col).Value = "Correct Utilty Expense"
ObjExcel.Cells(2, rept_correct_total_shel_cost_col).Value = "Correct Total Shelter Costs"
ObjExcel.Cells(2, rept_correct_net_adj_income_col).Value = "Correct Net Adj Income"
ObjExcel.Cells(2, rept_correct_hh_size_col).Value = "Correct HH Size"
ObjExcel.Cells(2, rept_snap_proration_col).Value = "Proration Date"
ObjExcel.Cells(2, rept_correct_snap_benefit_col).Value = "Correct SNAP Benefit"
ObjExcel.Cells(2, rept_orig_mf_caregivers_col).Value = "Orig Caregivers"
ObjExcel.Cells(2, rept_orig_mf_children_col).Value = "Orig Children"
ObjExcel.Cells(2, rept_orig_mf_earned_income_col).Value = "Orig MF Earned Income"
ObjExcel.Cells(2, rept_orig_mf_ei_deductions_col).Value = "Orig EI Disregards/Deductions"
ObjExcel.Cells(2, rept_orig_mf_net_ei_col).Value = "Orig Net Earned Income"
ObjExcel.Cells(2, rept_orig_mf_family_wage_level_col).Value = "Orig Family Wage Level"
ObjExcel.Cells(2, rept_orig_mf_difference_col).Value = "Orig Difference"
ObjExcel.Cells(2, rept_orig_mf_transitional_stndrd_col).Value = "Orig Trasitional Standard"
ObjExcel.Cells(2, rept_orig_mf_diff_or_trns_stndrd_col).Value = "Difference or Transitional Standard"
ObjExcel.Cells(2, rept_orig_mf_unearned_income_col).Value = "Orig MF Unearned Income"
ObjExcel.Cells(2, rept_orig_mf_unea_deductions_col).Value = "Orig Ded/Disrgd"
ObjExcel.Cells(2, rept_orig_mf_deemed_income_col).Value = "Orig Deemed Income"
ObjExcel.Cells(2, rept_orig_mf_cs_exclusion_col).Value = "Orig CS Exclusion"
ObjExcel.Cells(2, rept_orig_mf_subsidy_col).Value = "Orig Subsidy"
ObjExcel.Cells(2, rept_orig_mf_unmet_need_col).Value = "Orig Unmet Need"
ObjExcel.Cells(2, rept_orig_mf_mf_amt_col).Value = "Orig MF-MF"
ObjExcel.Cells(2, rept_orig_mf_fs_amt_col).Value = "Orig MF-FS"
ObjExcel.Cells(2, rept_orig_mf_hg_amt_col).Value = "Orig MF-HG"
ObjExcel.Cells(2, rept_correct_mf_caregivers_col).Value = "Correct Caregivers"
ObjExcel.Cells(2, rept_correct_mf_children_col).Value = "Correct Children"
ObjExcel.Cells(2, rept_correct_mf_earned_income_col).Value = "Correct MF Earned Income"
ObjExcel.Cells(2, rept_correct_mf_ei_deductions_col).Value = "Correct EI Disregards/Deductions"
ObjExcel.Cells(2, rept_correct_mf_net_ei_col).Value = "Correct Net Earned Income"
ObjExcel.Cells(2, rept_correct_mf_family_wage_level_col).Value = "Correct Family Wage Level"
ObjExcel.Cells(2, rept_correct_mf_difference_col).Value = "Correct Difference"
ObjExcel.Cells(2, rept_correct_mf_transitional_stndrd_col).Value = "Correct Trasitional Standard"
ObjExcel.Cells(2, rept_correct_mf_diff_or_trns_stndrd_col).Value = "Correct or Transitional Standard"
ObjExcel.Cells(2, rept_correct_mf_unearned_income_col).Value = "Correct MF Unearned Income"
ObjExcel.Cells(2, rept_correct_mf_unea_deductions_col).Value = "Correct Ded/Disrgd"
ObjExcel.Cells(2, rept_correct_mf_deemed_income_col).Value = "Correct Deemed Income"
ObjExcel.Cells(2, rept_correct_mf_cs_exclusion_col).Value = "Correct CS Exclusion"
ObjExcel.Cells(2, rept_correct_mf_subsidy_col).Value = "Correct Subsidy"
ObjExcel.Cells(2, rept_mf_proration_date_col).Value = "MFIP Proration Date"
ObjExcel.Cells(2, rept_correct_mf_unmet_need_col).Value = "Correct Unmet Need"
ObjExcel.Cells(2, rept_correct_mf_mf_amt_col).Value = "Correct MF-MF"
ObjExcel.Cells(2, rept_correct_mf_fs_amt_col).Value = "Correct MF-FS"
ObjExcel.Cells(2, rept_correct_mf_hg_amt_col).Value = "Correct MF-HG"
ObjExcel.Cells(2, rept_snap_pdf_link_col).Value = "PDF Calculation Sheet"
ObjExcel.Cells(2, rept_mfip_pdf_link_col).Value = "MFIP PDF Calculation Sheet"


objExcel.Rows(2).Font.Bold = True


excel_assignment_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\Yeng's Assignment List.xlsx"
Call excel_open(excel_assignment_file_path, False, False, ObjAssignExcel, objAssignWorkbook)

excel_row = 3
assign_row = 2
Do
    ObjExcel.Cells(excel_row, assign_worker_col).Value = ObjAssignExcel.Cells(assign_row, assign_worker_col).Value
    ObjExcel.Cells(excel_row, assign_case_numb_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value
    ObjExcel.Cells(excel_row, assign_case_name_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_name_col).Value
    ObjExcel.Cells(excel_row, assign_cash_col).Value = ObjAssignExcel.Cells(assign_row, assign_cash_col).Value
    ObjExcel.Cells(excel_row, assign_snap_col).Value = ObjAssignExcel.Cells(assign_row, assign_snap_col).Value
    ObjExcel.Cells(excel_row, assign_op_col).Value = ObjAssignExcel.Cells(assign_row, assign_op_col).Value
    ObjExcel.Cells(excel_row, assign_no_revw_form_col).Value = ObjAssignExcel.Cells(assign_row, assign_no_revw_form_col).Value
    ObjExcel.Cells(excel_row, assign_supp_col).Value = ObjAssignExcel.Cells(assign_row, assign_supp_col).Value
    ObjExcel.Cells(excel_row, assign_case_correct_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_correct_col).Value
    ObjExcel.Cells(excel_row, assign_notes_comments_col).Value = ObjAssignExcel.Cells(assign_row, assign_notes_comments_col).Value
    ObjExcel.Cells(excel_row, assign_tier_two_notes_col).Value = ObjAssignExcel.Cells(assign_row, assign_tier_two_notes_col).Value
    ObjExcel.Cells(excel_row, assign_tier_two_worker_col).Value = "Yeng"
    ObjExcel.Cells(excel_row, assign_tier_two_process_col).Value = "REVW"

    For array_counter = 0 to UBound(TIER_TWO_REVIEW_DETAIL_ARRAY, 2)
        ' MsgBox "TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter) - " & TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter) & vbCr &_
        '         "ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value - " & ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value
        If trim(TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter)) = trim(ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value) Then

            ObjExcel.Cells(excel_row, rept_case_numb_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter)
            ObjExcel.Cells(excel_row, rept_process_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(feb_process_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(fed_benefit_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(state_benefit_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_F_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_S_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_form_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_const, array_counter)
            ObjExcel.Cells(excel_row, rept_form_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_intv_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(interview_information_const, array_counter)
            ObjExcel.Cells(excel_row, rept_intv_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(interview_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_verifs_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(verifs_received_const, array_counter)
            ObjExcel.Cells(excel_row, rept_process_complete_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(process_complete_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_op_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_op_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_overpayment_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_overpayment_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_overpayment_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_overpayment_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_supp_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_supp_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_supplement_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_supplement_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_supplement_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_supplement_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_ded_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_net_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_housing_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_utility_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_shel_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_net_adj_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_hh_size_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(budgeted_hh_size_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_snap_benefit_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_ded_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_net_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_housing_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_utility_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_shel_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_net_adj_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_hh_size_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_hh_size_const, array_counter)
            ObjExcel.Cells(excel_row, rept_snap_proration_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_proration_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_snap_benefit_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_correct_amt_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_caregivers_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_caregivers_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_children_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_children_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_gross_total_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_ei_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_earned_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_net_ei_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_family_wage_level_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_difference_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_diff_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_transitional_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_ts_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_diff_or_trns_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_diff_or_ts_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_unea_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_unearned_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_deemed_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(deemed_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_unearned_income_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_cs_exclusion_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(cses_exclusion_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_unmet_need_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_total_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_subsidy_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_mf_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_fs_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_hg_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_caregivers_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_caregiver_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_children_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_children_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_ei_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_deductions_and_disreagards_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_net_ei_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_net_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_family_wage_level_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(familY_wage_level_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_difference_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(wage_level_difference_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_transitional_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(full_mfip_standard_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_diff_or_trns_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(difference_or_transitional_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unea_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_deductions_and_disreagards_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_deemed_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_deemed_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_cs_exclusion_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_cses_exclusion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_subsidy_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_mf_proration_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_proration_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unmet_need_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(prorated_unmet_need_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_mf_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_cash_portion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_fs_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_food_portion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_hg_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_hg_portion_const, array_counter)

            ' If snap_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, rept_snap_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & snap_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_snap_file_name & chr(34) & ")"
            ObjExcel.Cells(excel_row, rept_snap_pdf_link_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_pdf_excel_cell_info_const, array_counter)
            ' If mfip_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, rept_mfip_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & mfip_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_mfip_file_name & chr(34) & ")"
            ObjExcel.Cells(excel_row, rept_mfip_pdf_link_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_pdf_excel_cell_info_const, array_counter)

            Exit For
        End If
    Next


    excel_row = excel_row + 1
    assign_row = assign_row + 1
    next_case_numb = trim(ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value)
Loop until next_case_numb = ""

ObjAssignExcel.ActiveWorkbook.Close

ObjAssignExcel.Application.Quit
ObjAssignExcel.Quit

excel_assignment_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\Mandora's Assignment List.xlsx"
Call excel_open(excel_assignment_file_path, False, False, ObjAssignExcel, objAssignWorkbook)

ObjAssignExcel.worksheets("REVW").Activate

assign_row = 2
Do
    ObjExcel.Cells(excel_row, assign_worker_col).Value = ObjAssignExcel.Cells(assign_row, assign_worker_col).Value
    ObjExcel.Cells(excel_row, assign_case_numb_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value
    ObjExcel.Cells(excel_row, assign_case_name_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_name_col).Value
    ObjExcel.Cells(excel_row, assign_cash_col).Value = ObjAssignExcel.Cells(assign_row, assign_cash_col).Value
    ObjExcel.Cells(excel_row, assign_snap_col).Value = ObjAssignExcel.Cells(assign_row, assign_snap_col).Value
    ObjExcel.Cells(excel_row, assign_op_col).Value = ObjAssignExcel.Cells(assign_row, assign_op_col).Value
    ObjExcel.Cells(excel_row, assign_no_revw_form_col).Value = ObjAssignExcel.Cells(assign_row, assign_no_revw_form_col).Value
    ObjExcel.Cells(excel_row, assign_supp_col).Value = ObjAssignExcel.Cells(assign_row, assign_supp_col).Value
    ObjExcel.Cells(excel_row, assign_case_correct_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_correct_col).Value
    ObjExcel.Cells(excel_row, assign_notes_comments_col).Value = ObjAssignExcel.Cells(assign_row, assign_notes_comments_col).Value
    ObjExcel.Cells(excel_row, assign_tier_two_notes_col).Value = ObjAssignExcel.Cells(assign_row, assign_tier_two_notes_col).Value
    ObjExcel.Cells(excel_row, assign_tier_two_worker_col).Value = "Mandora"
    ObjExcel.Cells(excel_row, assign_tier_two_process_col).Value = "REVW"

    For array_counter = 0 to UBound(TIER_TWO_REVIEW_DETAIL_ARRAY, 2)
        ' MsgBox "TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter) - " & TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter) & vbCr &_
        '         "ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value - " & ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value
        If trim(TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter)) = trim(ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value) Then

            ObjExcel.Cells(excel_row, rept_case_numb_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter)
            ObjExcel.Cells(excel_row, rept_process_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(feb_process_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(fed_benefit_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(state_benefit_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_F_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_S_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_form_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_const, array_counter)
            ObjExcel.Cells(excel_row, rept_form_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_intv_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(interview_information_const, array_counter)
            ObjExcel.Cells(excel_row, rept_intv_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(interview_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_verifs_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(verifs_received_const, array_counter)
            ObjExcel.Cells(excel_row, rept_process_complete_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(process_complete_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_op_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_op_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_overpayment_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_overpayment_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_overpayment_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_overpayment_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_supp_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_supp_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_supplement_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_supplement_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_supplement_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_supplement_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_ded_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_net_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_housing_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_utility_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_shel_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_net_adj_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_hh_size_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(budgeted_hh_size_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_snap_benefit_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_ded_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_net_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_housing_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_utility_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_shel_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_net_adj_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_hh_size_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_hh_size_const, array_counter)
            ObjExcel.Cells(excel_row, rept_snap_proration_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_proration_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_snap_benefit_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_correct_amt_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_caregivers_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_caregivers_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_children_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_children_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_gross_total_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_ei_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_earned_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_net_ei_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_family_wage_level_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_difference_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_diff_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_transitional_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_ts_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_diff_or_trns_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_diff_or_ts_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_unea_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_unearned_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_deemed_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(deemed_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_unearned_income_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_cs_exclusion_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(cses_exclusion_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_unmet_need_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_total_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_subsidy_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_mf_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_fs_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_hg_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_caregivers_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_caregiver_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_children_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_children_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_ei_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_deductions_and_disreagards_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_net_ei_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_net_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_family_wage_level_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(familY_wage_level_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_difference_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(wage_level_difference_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_transitional_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(full_mfip_standard_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_diff_or_trns_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(difference_or_transitional_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unea_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_deductions_and_disreagards_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_deemed_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_deemed_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_cs_exclusion_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_cses_exclusion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_subsidy_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_mf_proration_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_proration_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unmet_need_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(prorated_unmet_need_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_mf_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_cash_portion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_fs_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_food_portion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_hg_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_hg_portion_const, array_counter)

            ' If snap_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_snap_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & snap_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_snap_file_name & chr(34) & ")"
            ObjExcel.Cells(excel_row, rept_snap_pdf_link_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_pdf_excel_cell_info_const, array_counter)
            ' If mfip_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_mfip_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & mfip_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_mfip_file_name & chr(34) & ")"
            ObjExcel.Cells(excel_row, rept_mfip_pdf_link_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_pdf_excel_cell_info_const, array_counter)

            Exit For
        End If
    Next


    excel_row = excel_row + 1
    assign_row = assign_row + 1
    next_case_numb = trim(ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value)
Loop until next_case_numb = ""


ObjAssignExcel.worksheets("MONT").Activate

assign_row = 2
Do
    ObjExcel.Cells(excel_row, assign_worker_col).Value = ObjAssignExcel.Cells(assign_row, assign_worker_col).Value
    ObjExcel.Cells(excel_row, assign_case_numb_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value
    ObjExcel.Cells(excel_row, assign_case_name_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_name_col).Value
    ObjExcel.Cells(excel_row, assign_cash_col).Value = ObjAssignExcel.Cells(assign_row, assign_cash_col).Value
    ObjExcel.Cells(excel_row, assign_snap_col).Value = ObjAssignExcel.Cells(assign_row, assign_snap_col).Value
    ObjExcel.Cells(excel_row, assign_op_col).Value = ObjAssignExcel.Cells(assign_row, assign_op_col).Value
    ObjExcel.Cells(excel_row, assign_no_revw_form_col).Value = ObjAssignExcel.Cells(assign_row, assign_no_revw_form_col).Value
    ObjExcel.Cells(excel_row, assign_supp_col).Value = ObjAssignExcel.Cells(assign_row, assign_supp_col).Value
    ObjExcel.Cells(excel_row, assign_case_correct_col).Value = ObjAssignExcel.Cells(assign_row, assign_case_correct_col).Value
    ObjExcel.Cells(excel_row, assign_notes_comments_col).Value = ObjAssignExcel.Cells(assign_row, assign_notes_comments_col).Value
    ObjExcel.Cells(excel_row, assign_tier_two_notes_col).Value = ObjAssignExcel.Cells(assign_row, assign_tier_two_notes_col).Value
    ObjExcel.Cells(excel_row, assign_tier_two_worker_col).Value = "Mandora"
    ObjExcel.Cells(excel_row, assign_tier_two_process_col).Value = "MONT"

    For array_counter = 0 to UBound(TIER_TWO_REVIEW_DETAIL_ARRAY, 2)
        If trim(TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter)) = trim(ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value) Then

            ObjExcel.Cells(excel_row, rept_case_numb_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(case_number_const, array_counter)
            ObjExcel.Cells(excel_row, rept_process_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(feb_process_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(fed_benefit_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(state_benefit_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_F_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_S_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_issued_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_form_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_const, array_counter)
            ObjExcel.Cells(excel_row, rept_form_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(form_received_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_intv_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(interview_information_const, array_counter)
            ObjExcel.Cells(excel_row, rept_intv_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(interview_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_verifs_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(verifs_received_const, array_counter)
            ObjExcel.Cells(excel_row, rept_process_complete_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(process_complete_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_op_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_op_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_overpayment_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_overpayment_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_overpayment_const, array_counter)
            ObjExcel.Cells(excel_row, rept_op_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_overpayment_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_fed_supp_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(SNAP_state_supp_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_mf_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_cash_supplement_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_fs_s_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_s_supplement_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_fs_f_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_food_f_supplement_const, array_counter)
            ObjExcel.Cells(excel_row, rept_supp_mf_hg_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_hg_supplement_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_ded_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_net_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_housing_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_utility_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_total_shel_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_net_adj_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_hh_size_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(budgeted_hh_size_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_snap_benefit_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(earned_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(unearned_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_ded_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_deduction_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_net_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_housing_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_housing_cost_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_utility_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(utilities_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_total_shel_cost_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_shelter_cost_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_net_adj_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(net_adj_income_correct_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_hh_size_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_hh_size_const, array_counter)
            ObjExcel.Cells(excel_row, rept_snap_proration_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_proration_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_snap_benefit_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_correct_amt_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_caregivers_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_caregivers_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_children_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_children_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_gross_total_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_ei_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_earned_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_net_ei_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_family_wage_level_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_difference_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_fwl_diff_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_transitional_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_ts_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_diff_or_trns_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mf_orig_diff_or_ts_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_unea_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_orig_deductions_unearned_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_deemed_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(deemed_income_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_budgeted_unearned_income_const, array_counter)

            ObjExcel.Cells(excel_row, rept_orig_mf_cs_exclusion_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(cses_exclusion_budgeted_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_unmet_need_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_total_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_subsidy_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_mf_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_MF_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_fs_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_FS_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_orig_mf_hg_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_MF_HG_issued_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_caregivers_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_caregiver_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_children_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_children_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_earned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_ei_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_earned_deductions_and_disreagards_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_net_ei_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_net_earned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_family_wage_level_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(familY_wage_level_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_difference_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(wage_level_difference_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_transitional_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(full_mfip_standard_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_diff_or_trns_stndrd_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(difference_or_transitional_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unearned_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_income_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unea_deductions_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(total_correct_mfip_unearned_deductions_and_disreagards_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_deemed_income_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_deemed_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_cs_exclusion_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(correct_mfip_cses_exclusion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_subsidy_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_subsidy_tribal_amt_const, array_counter)
            ObjExcel.Cells(excel_row, rept_mf_proration_date_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_proration_date_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_unmet_need_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(prorated_unmet_need_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_mf_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_cash_portion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_fs_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_food_portion_const, array_counter)
            ObjExcel.Cells(excel_row, rept_correct_mf_hg_amt_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_correct_hg_portion_const, array_counter)

            ' If snap_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_snap_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & snap_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_snap_file_name & chr(34) & ")"
            ObjExcel.Cells(excel_row, rept_snap_pdf_link_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(snap_pdf_excel_cell_info_const, array_counter)
            ' If mfip_pdf_file_save_path <> "" Then ObjDetailsExcel.Cells(total_excel_row, det_mfip_pdf_link_col).Value 				= "=HYPERLINK(" & chr(34) & mfip_pdf_file_save_path & chr(34) & ", " & chr(34) & pdf_mfip_file_name & chr(34) & ")"
            ObjExcel.Cells(excel_row, rept_mfip_pdf_link_col).Value = TIER_TWO_REVIEW_DETAIL_ARRAY(mfip_pdf_excel_cell_info_const, array_counter)

            Exit For
        End If
    Next


    excel_row = excel_row + 1
    assign_row = assign_row + 1
    next_case_numb = trim(ObjAssignExcel.Cells(assign_row, assign_case_numb_col).Value)
Loop until next_case_numb = ""


ObjAssignExcel.ActiveWorkbook.Close

ObjAssignExcel.Application.Quit
ObjAssignExcel.Quit

MsgBox "STOP HERE"

'read the BOBI List
ObjExcel.worksheets("BOBI List").Activate

excel_row = 2
array_item = 0
Do
    bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
    If InStr(bobi_cases_string, "~" & bobi_case_numb & "~") = 0 Then
        ReDim Preserve BOBI_LIST_ARRAY(last_const, array_item)
        BOBI_LIST_ARRAY(case_numb, array_item) = bobi_case_numb
        BOBI_LIST_ARRAY(app_date, array_item) = ObjExcel.Cells(excel_row, 6).Value

        array_item = array_item + 1
    End If
    excel_row = excel_row + 1
    next_bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
Loop until next_bobi_case_numb = ""

'read the SQL list
ObjExcel.worksheets("SQL List").Activate

excel_row = 2
array_item = 0
Do
    bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
    If InStr(bobi_cases_string, "~" & bobi_case_numb & "~") = 0 Then
        ReDim Preserve SQL_LIST_ARRAY(last_const, array_item)
        SQL_LIST_ARRAY(case_numb, array_item) = bobi_case_numb
        SQL_LIST_ARRAY(app_date, array_item) = ObjExcel.Cells(excel_row, 4).Value
        SQL_LIST_ARRAY(found, array_item) = False

        array_item = array_item + 1
    End If
    excel_row = excel_row + 1
    next_bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
Loop until next_bobi_case_numb = ""





'loop through the BOBI list add the SQL
ObjExcel.worksheets("Data Compare").Activate

excel_row = 2
For bobi_item = 0 to UBound(BOBI_LIST_ARRAY, 2)
    MAXIS_case_number = BOBI_LIST_ARRAY(case_numb, bobi_item)
    ObjExcel.Cells(excel_row, 1).Value = BOBI_LIST_ARRAY(case_numb, bobi_item)
    ObjExcel.Cells(excel_row, 2).Value = True
    ObjExcel.Cells(excel_row, 3).Value = BOBI_LIST_ARRAY(app_date, bobi_item)
    ObjExcel.Cells(excel_row, 4).Value = False
    For sql_item = 0 to UBound(SQL_LIST_ARRAY, 2)
        If SQL_LIST_ARRAY(case_numb, sql_item) = BOBI_LIST_ARRAY(case_numb, bobi_item) Then
            SQL_LIST_ARRAY(found, sql_item) = True
            ObjExcel.Cells(excel_row, 4).Value = True
            ObjExcel.Cells(excel_row, 5).Value = SQL_LIST_ARRAY(app_date, sql_item)
            Exit For
        End If
    Next
    Call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash_1_status, 4, 6, 74
    EMReadScreen cash_1_intvw, 8, 6, 55
    EMReadScreen cash_2_status, 4, 7, 74
    EMReadScreen cash_2_intvw, 8, 7, 55
    EMReadScreen snap_status, 4, 10, 74
    EMReadScreen snap_intvw, 8, 10, 55

    cash_1_intvw = replace(cash_1_intvw, " ", "/")
    If cash_1_intvw = "__/__/__" Then cash_1_intvw = ""
    cash_2_intvw = replace(cash_2_intvw, " ", "/")
    If cash_2_intvw = "__/__/__" Then cash_2_intvw = ""
    snap_intvw = replace(snap_intvw, " ", "/")
    If snap_intvw = "__/__/__" Then snap_intvw = ""

    ObjExcel.Cells(excel_row, 6).Value = cash_1_status
    ObjExcel.Cells(excel_row, 7).Value = cash_1_intvw
    ObjExcel.Cells(excel_row, 8).Value = cash_2_status
    ObjExcel.Cells(excel_row, 9).Value = cash_2_intvw
    ObjExcel.Cells(excel_row, 10).Value = snap_status
    ObjExcel.Cells(excel_row, 11).Value = snap_intvw

    Call back_to_SELF

    excel_row = excel_row + 1
Next

For sql_item = 0 to UBound(SQL_LIST_ARRAY, 2)
    If SQL_LIST_ARRAY(found, sql_item) = False Then
        MAXIS_case_number = SQL_LIST_ARRAY(case_numb, sql_item)
        ObjExcel.Cells(excel_row, 1).Value = SQL_LIST_ARRAY(case_numb, sql_item)
        ObjExcel.Cells(excel_row, 2).Value = False
        ObjExcel.Cells(excel_row, 3).Value = ""
        ObjExcel.Cells(excel_row, 4).Value = True
        ObjExcel.Cells(excel_row, 5).Value = SQL_LIST_ARRAY(app_date, sql_item)

        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen cash_1_status, 4, 6, 74
        EMReadScreen cash_1_intvw, 8, 6, 55
        EMReadScreen cash_2_status, 4, 7, 74
        EMReadScreen cash_2_intvw, 8, 7, 55
        EMReadScreen snap_status, 4, 10, 74
        EMReadScreen snap_intvw, 8, 10, 55

        cash_1_intvw = replace(cash_1_intvw, " ", "/")
        If cash_1_intvw = "__/__/__" Then cash_1_intvw = ""
        cash_2_intvw = replace(cash_2_intvw, " ", "/")
        If cash_2_intvw = "__/__/__" Then cash_2_intvw = ""
        snap_intvw = replace(snap_intvw, " ", "/")
        If snap_intvw = "__/__/__" Then snap_intvw = ""

        ObjExcel.Cells(excel_row, 6).Value = cash_1_status
        ObjExcel.Cells(excel_row, 7).Value = cash_1_intvw
        ObjExcel.Cells(excel_row, 8).Value = cash_2_status
        ObjExcel.Cells(excel_row, 9).Value = cash_2_intvw
        ObjExcel.Cells(excel_row, 10).Value = snap_status
        ObjExcel.Cells(excel_row, 11).Value = snap_intvw

        Call back_to_SELF

        excel_row = excel_row + 1
    End If
Next

call script_end_procedure("Done")
