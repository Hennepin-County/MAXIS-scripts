'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - EARNED INCOME BUDGETING.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
Call changelog_update("10/07/2025", "NEW SCRIPT VERSION - MAJOR UPDATE##~## ##~##Primary update is to allow for budgeting details of CASH to align with the process for the SNAP budgeting.##~## - Check details now entered on individual dialogs.##~## - Exclusions can be indicated on any check for SNAP, CASH, or both.##~## - Some updates to the format of the CASE/NOTE.##~## - Base functionality should all remain.##~## ##~##Interactions with this script will take some time to become familiar - additional information is available in the instructions (look for the button on the main dialog).", "Casey Love, Hennepin County")
Call changelog_update("04/18/2025", "Fixed issue with dialog sizing for entered checks.", "Mark Riegel, Hennepin County")
Call changelog_update("03/31/2025", "Updated script to display conversation/clarification field regardless of program(s).", "Mark Riegel, Hennepin County")
Call changelog_update("02/01/2025", "Support for MFIP, GA, and UHFS Budgeting Workaround to support the policy change to eliminate Monthly Reporting and Retrospective Budgeting.##~## ##~##These updates follow the Guide to Six-Mount Budgeting available in SIR. Details of how the JOBS panel is updated can be found in this guide.##~## ##~##As with any new functionality, but particularly when the supporting policy is also new, reach out with any questions or script errors.##~##", "Casey Love, Hennepin County")
Call changelog_update("05/23/2024", "BUG FIX - Ensuring the case can be updated in STAT if run in production. This will prevent errors in panels not updating on inactive or out of county cases.", "Casey Love, Hennepin County")
Call changelog_update("03/14/2024", "BUG FIX - Error when trying to find the correct JOBS panel in instances where the job name changes. This should now provide an option to select the correct panel.", "Casey Love, Hennepin County")
Call changelog_update("07/26/2023", "Multiple updates to the script##~####~##ENHANCEMENTS: ##~##-Added 'Save your Work' functionality.##~##-Added an option to identify a check is a 'Bonus Check'.##~##-Option to break out pay amount in different types.##~##-Added a YTD Calculator.##~####~##BUG FIXED: ##~##-If more than 5 unique check dates are entered for a month, they will be entered as a lump payment.##~##-If multiple checks are entered for the same date, JOBS panel will be updated as a single payment.##~##-Better handling for a job for a new HH member by checking the arrival date.##~##-Allow for worker to select the correct JOBS panel to update if the script cannot find it.##~####~##There is a lot of new functionality and this is a fairly complicated script run. Please contact the script team if you have any concerns or questions.##~##", "Casey Love, Hennepin County")
Call changelog_update("03/25/2021", "Added information buttons to the dialogs.##~## ##~##There are a number of new buttons with a '!' on it that will display some tips about the policy and use of these dialogs. Click on them to find out more.##~## ##~##There are also direct links to the instruction documents on SharePoint.##~##", "Casey Love, Hennepin County")
Call changelog_update("09/17/2020", "Update to the script to remove the functionality that would LUMP together any income in the month of application on the SNAP PIC.##~## ##~##The income will still update the SNAP PIC with income as a LUMP if for the month a job started.", "Casey Love, Hennepin County")
Call changelog_update("09/17/2020", "The script will no longer read the Pay Frequency. This will have to be entered when entering the paycheck information.##~## ##~##We made this change because this script relys heavily on the pay frequency being correct at this point and there is not a great way to ensure accuracy otherwise.", "Casey Love, Hennepin County")
Call changelog_update("09/17/2020", "Added some options to the 'Explanation of why we are not using 30 days of income' for SNAP. This used to be a typing field (EditBox) and now had a dropdown option but you can still type anything that explains this information. ##~## ##~##For SNAP anytime we use anything other than 30 days of income, we are required to note clearly why we used something other than 30 day of income. Now there are some common options listed.##~## ##~##If you have more options that happen regularly, please send them to us for review and we can possibly add them. Remember you can always type out the explanation as well.", "Casey Love, Hennepin County")
Call changelog_update("01/10/2020", "This script may force an error report at the end of the script run.##~## ##~##We have and ongoing script error that is happening around dates and the updating of the panel. We have added a workaround to the script but if this error happens, it should be sending us and error report so we can try to discover the nature of the issue in real time. ##~## The script should be working well and this is just an alert that the error reporting may happen automatically. ", "Casey Love, Hennepin County")
Call changelog_update("01/08/2020", "The script cannot be used on a JOBS panel with an income end date at this time. The script now reads if an end date exists and prevents information from being entered for a panel that cannot be updated. ##~##", "Casey Love, Hennepin County")
call changelog_update("1/3/2020", "BUG FIX - The script could not continue if ongoing income for a job was $0. Updated functionality to better suppor an ongoing job with $0 income. This new functionality still does not support jobs that are ending.##~##", "Casey Love, Hennepin County")
call changelog_update("12/16/2019", "BUG FIX - There was an error when completing the PIC in the month of application. This should now be resolved and the script will not get stuck on the PIC.##~##", "Casey Love, Hennepin County")
call changelog_update("11/22/2019", "BUG FIX - The PIC does not allow for hours to have more than 2 decimal points written into MAXIS. Sometimes check stubs have 3 decimals provided. The script will change to 2 decimal points for the entry of information only, the information entered into the dialog and input on the CASE/NOTE can still be 3 decimal points.##~##", "Casey Love, Hennepin County")
call changelog_update("11/06/2019", "BUG FIX - The script was hitting an error if a 'known pay date' was entered that is after the intital month to update. Added functionality for the script to recalculate the 'known pay date' back to the beginning of the update period. This way any known pay date will work in the script.##~##", "Casey Love, Hennepin County")
call changelog_update("08/21/2019", "Handling added to prohibit the attempted update of months prior to the first check entered for a job.", "Casey Love, Hennepin County")
call changelog_update("08/13/2019", "Bug fix when updating RETRO checks on occasion that causes the script to stop and error.", "Casey Love, Hennepin County")
call changelog_update("08/05/2019", "Bug fix in script that would sometimes enter the wrong dates on the RETRO side if checks that were out of schedule are used. Additionally, added functionality to find the weekday of pay to better determine which check is out of schedule.", "Casey Love, Hennepin County")
call changelog_update("08/02/2019", "Bug in script when 2 checks with the same date are entered, script would get stuck and be unable to continue. Script will now continue but confirmation of the paydate will be required as the script reads it as an unexpected pay date.", "Casey Love, Hennepin County")
call changelog_update("07/31/2019", "Bug fix where occasionally the script fails at navigating to the JOBS panel for update.", "Casey Love, Hennepin County")
call changelog_update("06/27/2019", "Bug fix on Case Noting Retro HC Income.", "Casey Love, Hennepin County")
Call changelog_update("06/13/2019", "Bug fix for semi-monthly pay frequency when no actual checks are listed (only anticipated income schedule and hours).", "Casey Love, Hennepin County")
Call changelog_update("06/05/2019", "Bug fix for SNAP cases that are currently set to close, a date error was preventing the script from running.", "Casey Love, Hennepin County")
Call changelog_update("04/24/2019", "Added wording to the Confirm Budget Dialog that explains the functionality of the script. The script will return to the enter pay dialog if the budgets are not indicated as correct on the confirm budget dialog. This functionality is not new, it was built to go back when the budget is not confirmed.", "Casey Love, Hennepin County")
Call changelog_update("03/26/2019", "Fixed errors when pay is twice monthly. Added better handling for reading the employer name.", "Casey Love, Hennepin County")
call changelog_update("03/05/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS===============================================================================================================
class jobs_income
	'PANEL INFO
	public new_panel
	public member
	public instance
	public employer
	public employer_with_underscores
	public job_type
	public verif_type
	public hourly_wage
	public snap_hourly_wage
	public cash_hourly_wage
	public income_start_dt
	public income_end_dt
	public pay_freq
	public updated_date
	public old_verif
	public first_retro_check
	public initial_month_mo
	public initial_month_yr
    public RETRO_month
    public RETRO_footer_month
    public RETRO_footer_year

	'DIALOG INFO
	public hrs_per_wk
	public pay_per_hr
	public display_hrs_per_wk
	public display_pay_per_hr
	public known_pay_date
	public word_for_freq
	public apply_to_SNAP
	public apply_to_CASH
	public apply_to_HC
	public apply_to_GRH
    public prog_list
	public verif_date
	public verif_explain
	public selection_rsn
	public income_excluded_cash_reason
	public hc_budg_notes
	public hc_retro
	public EI_panel_vbYes
	public spoke_with
	public convo_detail
	public paycheck_list_title
	public excl_cash_rsn

	'LOGIC DETAILS
    public budget_confirmed
	public SNAP_list_of_excluded_pay_dates
	public CASH_list_of_excluded_pay_dates
	public pay_weekday
	public income_received
	public all_pay_in_app_month
	public there_are_counted_checks
	public actual_checks_provided
	public anticipated_income_provided
	public missing_checks_list
	public issues_with_frequency
	public default_start_date
	public bimonthly_first
	public bimonthly_second
	public ignore_antic
	' public antic_pay_list
	public pick_one
	public days_to_add
	public months_to_add

	public first_check
	public last_check
	public order_ubound
	public checks_exist
	public estimate_exists
    public gross_max_string_len

	public total_check_count
	public total_gross_amount
	public total_hours
	public ave_hrs_per_pay
	public ave_inc_per_pay
	public monthly_income

	public snap_check_count
	public snap_budgeted_total
	public snap_hours
	public snap_ave_hrs_per_pay
	public snap_ave_inc_per_pay
	public snap_hrs_per_wk
	public SNAP_monthly_income
	public snap_anticipated_pay_array

	public cash_check_count
	public cash_budgeted_total
	public cash_hours
	public cash_ave_hrs_per_pay
	public cash_ave_inc_per_pay
	public cash_hrs_per_wk
	public CASH_monthly_income
	public cash_anticipated_pay_array

	'CURRENTLY UNUSED
	public numb_months
	public reg_non_monthly

	'UPDATE FUNCTIONALITY
	public update_this_month
	public months_updated
	public income_lumped_mo
	public lump_reason
	public act_checks_lumped
	public est_checks_lumped
	public lump_gross
	public lump_hrs
	public mo_w_more_5_chcks
	public update_future
    public updates_to_display

	public next_check_btn
	public cancel_check_btn
	public save_details_btn
	public delete_check_btn


	public expected_check_array()
    public this_month_checks_array()
    public retro_month_checks_array()

    public cash_array_info_exists
    public cash_info_cash_mo_yr()
    public cash_info_retro_mo_yr()
    public cash_info_retro_updtd()
    public cash_info_prosp_updtd()
    public cash_info_mo_retro_pay()
    public cash_info_mo_retro_hrs()
    public cash_info_mo_prosp_pay()
    public cash_info_mo_prosp_hrs()

	public pay_date()
	public gross_amount()
	public hours()
	public exclude_entirely()
	public exclude_from_SNAP()
	public exclude_from_CASH()
	public reason_to_exclude()
	public exclude_ALL_amount()
    public exclude_ALL_hours()
	public exclude_SNAP_amount()
	public exclude_SNAP_hours()
	public exclude_CASH_amount()
	public exclude_CASH_hours()
	public SNAP_info_string()
	public CASH_info_string()
	public check_order()
	public view_pay_date()
	public frequency_issue()
	public future_check()
	public duplicate_pay_date()
	public reason_SNAP_amt_excluded()
	public reason_CASH_amt_excluded()
	public pay_detail_btn()
	public check_info_entered()
	public bonus_check()
	public pay_split_regular_amount()
	public pay_split_bonus_amount()
	public pay_split_ot_amount()
    public pay_split_ot_hours()
	public pay_split_shift_diff_amount()
	public pay_split_tips_amount()
	public pay_split_other_amount()
	public pay_split_other_detail()
	public pay_excld_bonus()
	public pay_excld_ot()
	public pay_excld_shift_diff()
	public pay_excld_tips()
	public pay_excld_other()
	public split_check_string()
    public split_check_excld_string()
    public split_exclude_amount()
	public duplct_pay_date()
	public calculated_by_ytd()
	public ytd_calc_notes()
	public pay_detail_exists()
	public combined_into_one()
	public SNAP_dialog_display()
	public CASH_dialog_display()


	public sub add_check()
        'records the date, amount and hours then calls the check details dialog to add a new check to the list of checks for the job

		check_count = UBound(pay_date)
		If pay_date(check_count) <> "" Then check_count = check_count + 1

		Do
			pay_date_fld = ""
			pay_amt_fld = ""
			hours_fld = ""
			bonus_checkbox = unchecked
			Do
				err_msg = ""
				save_check = False
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 246, 165,  "Enter Check Info"
                    Text 55, 10, 225, 10, "***  Add the basic check details here.  ***"
					GroupBox 10, 20, 225, 70, "Actual Check Details"
					Text 20, 40, 35, 10, "Pay Date:"
					EditBox 55, 35, 50, 15, pay_date_fld
					Text 115, 40, 60, 10, "Gross Amount: $"
					EditBox 175, 35, 50, 15, pay_amt_fld
					Text 125, 60, 45, 10, "Total Hours:"
					EditBox 175, 55, 50, 15, hours_fld
					Text 55, 50, 50, 10, "(MM/DD/YY)"
					CheckBox 55, 75, 205, 10, "Check Here if this is a BONUS CHECK", bonus_checkbox
					'TODO - maybe add a 'DONE' selection
                    Text 10, 95, 230, 30, "Once you save the initial check information here, a new dialog will open. Details including pay splits and exclusions can be entered on the next dialog. Remember that checks do NOT need to be entered in order."
					ButtonGroup ButtonPressed
						PushButton 145, 145, 90, 15, "Save and Add Details", next_check_btn
						PushButton 10, 130, 90, 15, "Cancel This Check", cancel_check_btn                       'This button exists in two different forms for clarity to users but really they both mean 'I'm done with this now'
						PushButton 10, 145, 130, 15, "Done Adding Checks / See Check List", cancel_check_btn
				EndDialog

				dialog Dialog1
                save_your_work

				If ButtonPressed = -1 Then ButtonPressed = next_check_btn
				If ButtonPressed = 0 Then ButtonPressed = cancel_check_btn

				If NOT IsDate(pay_date_fld) Then err_msg = err_msg & vbCr & "* Pay Date should be entered as a date."
				If IsDate(pay_date_fld) Then
                    in_appl_month = False
                    If IsDate(fs_appl_date) Then
                        If DatePart("m", fs_appl_date) = DatePart("m", pay_date_fld) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", pay_date_fld) Then in_appl_month = True
                    End If
                    If DateDiff("d", date, pay_date_fld) > 0 and NOT in_appl_month Then             'if the pay date is in the future we have to error
                        err_msg = err_msg & vbNewLine & "* Paydates cannot be in the future. (" & pay_date_fld & ")"
                    End If
				End If
				If NOT IsNumeric(pay_amt_fld) Then err_msg = err_msg & vbCr & "* The Amount of the Pay Check should be entered."
				If NOT IsNumeric(hours_fld) Then err_msg = err_msg & vbCr & "* The Hours worked should be entered as a number."
				If ButtonPressed = cancel_check_btn Then err_msg = ""

				If err_msg <> "" Then MsgBox "*  *  *  NOTICE  *  *  *" & vbCr & "Please Update the information in the dialog to continue:" & vbCr & err_msg
				If ButtonPressed = next_check_btn Then save_check = True
			Loop until err_msg = ""

			If save_check = True Then
				call resize_check_list(check_count)

				pay_date(check_count) = DateAdd("d", 0, pay_date_fld)
				gross_amount(check_count) = FormatNumber(pay_amt_fld, 2, -1, 0, 0)
				' pay_split_regular_amount(check_count) = pay_amt_fld
				hours(check_count) = hours_fld
				bonus_check(check_count) = False
				duplicate_pay_date(check_count) = False
				If bonus_checkbox = checked then
                    bonus_check(check_count) = True
                    exclude_entirely(check_count) = True
                    reason_to_exclude(check_count) = "Bonus Check is not expected to continue."
                End If
				pay_detail_btn(check_count) = 2000+check_count
				duplct_pay_date(check_count) = False
				check_info_entered(check_count) = True
				future_check(check_count) = False
                If IsDate(fs_appl_date) Then
                    If DatePart("m", fs_appl_date) = DatePart("m", pay_date(check_count)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", pay_date(check_count)) Then   'if the paydate is in the application month
                        If DateDiff("d", date, pay_date(check_count)) > 0 Then future_check(check_count) = TRUE   'this is a future check
                    End If
				End If

				call check_details_dialog(False, check_count, ButtonPressed)

				If ButtonPressed = delete_check_btn Then
					check_info_entered(check_count) = False
					check_count = check_count - 1
					call resize_check_list(check_count)

					check_count = check_count + 1
				End If
			End If
			If save_check = True Then check_count = check_count + 1
		Loop until ButtonPressed = cancel_check_btn
		Call evaluate_checks
		If checks_exist Then Call order_checks
	end sub

	public sub calculate_totals()
        'does a bunch of calculations with the checks or anticipated income to determine totals, averages, hours per week, etc.

        gross_max_string_len = 0
		If actual_checks_provided Then
			paycheck_list_title = "Paychecks Provided for Determination:"

			If IsNumeric(pay_per_hr) Then
				hourly_wage = pay_per_hr
                snap_hourly_wage = pay_per_hr
                cash_hourly_wage = pay_per_hr
			Else
				total_hourly_amount = 0
				snap_total_hourly_amount = 0
				cash_total_hourly_amount = 0
				total_hours = 0
				snap_total_hours = 0
				cash_total_hours = 0
				For all_income = 0 to UBound(pay_date)
					If NOT future_check(all_income) and NOT bonus_check(all_income) Then
						exclude_amount = 0
						If IsNumeric(pay_split_bonus_amount(all_income)) Then exclude_amount = exclude_amount + pay_split_bonus_amount(all_income)
						If IsNumeric(pay_split_shift_diff_amount(all_income)) Then exclude_amount = exclude_amount + pay_split_shift_diff_amount(all_income)
						If IsNumeric(pay_split_tips_amount(all_income)) Then exclude_amount = exclude_amount + pay_split_tips_amount(all_income)
						If IsNumeric(pay_split_other_amount(all_income)) Then exclude_amount = exclude_amount + pay_split_other_amount(all_income)
						If IsNumeric(pay_split_ot_amount(all_income)) and IsNumeric(pay_split_ot_hours(all_income)) Then exclude_amount = exclude_amount + pay_split_ot_amount(all_income)
						If IsNumeric(exclude_ALL_amount(all_income)) and IsNumeric(exclude_ALL_hours(all_income)) Then exclude_amount = exclude_amount + exclude_ALL_amount(all_income)

                        total_hourly_amount = total_hourly_amount + gross_amount(all_income)

                        exclude_hours = 0
						If IsNumeric(pay_split_ot_amount(all_income)) and IsNumeric(pay_split_ot_hours(all_income)) Then exclude_hours = exclude_hours + pay_split_ot_hours(all_income)
						If IsNumeric(exclude_ALL_amount(all_income)) and IsNumeric(exclude_ALL_hours(all_income)) Then exclude_hours = exclude_hours + exclude_ALL_hours(all_income)
						total_hours = total_hours + hours(all_income)

                        If exclude_from_SNAP(all_income) = unchecked Then
                            snap_total_hourly_amount = snap_total_hourly_amount + total_hourly_amount - exclude_amount
                            If IsNumeric(exclude_SNAP_amount(all_income)) and IsNumeric(exclude_SNAP_hours(all_income))Then snap_total_hourly_amount = snap_total_hourly_amount - exclude_SNAP_amount(all_income)
                            snap_total_hours = snap_total_hours + total_hours - exclude_hours
                            If IsNumeric(exclude_SNAP_amount(all_income)) and IsNumeric(exclude_SNAP_hours(all_income)) Then snap_total_hours = snap_total_hours - exclude_SNAP_hours(all_income)
                        End If

                        If exclude_from_CASH(all_income) = unchecked Then
                            cash_total_hourly_amount = cash_total_hourly_amount + total_hourly_amount - exclude_amount
                            If IsNumeric(exclude_CASH_amount(all_income)) and IsNumeric(exclude_CASH_hours(all_income)) Then cash_total_hourly_amount = cash_total_hourly_amount - exclude_CASH_amount(all_income)
                            cash_total_hours = cash_total_hours + total_hours - exclude_hours
                            If IsNumeric(exclude_CASH_amount(all_income)) and IsNumeric(exclude_CASH_hours(all_income)) Then cash_total_hours = cash_total_hours - exclude_CASH_hours(all_income)
                        End If
                    End If
				Next
				hourly_wage = total_hourly_amount/total_hours
				pay_per_hr = total_hourly_amount/total_hours
                snap_hourly_wage = snap_total_hourly_amount/snap_total_hours
                cash_hourly_wage = cash_total_hourly_amount/cash_total_hours

			End If
			hourly_wage = FormatNumber(hourly_wage, 2,,0)
			snap_hourly_wage = FormatNumber(snap_hourly_wage, 2,,0)
			cash_hourly_wage = FormatNumber(cash_hourly_wage, 2,,0)

			cash_check_count = 0
			snap_check_count = 0
			total_check_count = 0
			cash_budgeted_total = 0
			snap_budgeted_total = 0
			total_gross_amount = 0
			cash_hours = 0
			snap_hours = 0
			total_hours = 0
			SNAP_list_of_excluded_pay_dates = ""
			CASH_list_of_excluded_pay_dates = ""
			For order_number = 1 to order_ubound                        'loop through the order number lowest to highest

				For all_income = 0 to UBound(pay_date)
					If check_order(all_income) = order_number Then
                        gross_amt_len = Len(CStr(gross_amount(all_income)))
                        If gross_amt_len > gross_max_string_len Then gross_max_string_len = gross_amt_len

						split_exclude_amount(all_income) = 0
						SNAP_dialog_display(all_income) = ""
						CASH_dialog_display(all_income) = ""
						If NOT future_check(all_income) Then
							total_check_count = total_check_count + 1
							total_gross_amount = total_gross_amount + gross_amount(all_income)
							total_hours = total_hours + hours(all_income)

							If NOT exclude_entirely(all_income) Then
								hourly_amount = 0
								If pay_excld_bonus(all_income) = checked Then split_exclude_amount(all_income) = split_exclude_amount(all_income) + pay_split_bonus_amount(all_income)
								If pay_excld_ot(all_income) = checked Then split_exclude_amount(all_income) = split_exclude_amount(all_income) + pay_split_ot_amount(all_income)
								If pay_excld_shift_diff(all_income) = checked Then split_exclude_amount(all_income) = split_exclude_amount(all_income) + pay_split_shift_diff_amount(all_income)
								If pay_excld_tips(all_income) = checked Then split_exclude_amount(all_income) = split_exclude_amount(all_income) + pay_split_tips_amount(all_income)
								If pay_excld_other(all_income) = checked Then split_exclude_amount(all_income) = split_exclude_amount(all_income) + pay_split_other_amount(all_income)
								If pay_excld_ot(all_income) = checked Then hourly_amount = hourly_amount + pay_split_ot_hours(all_income)

								SNAP_dialog_display(all_income) = gross_amount(all_income) & " - " & hours(all_income) & " hrs."
								CASH_dialog_display(all_income) = gross_amount(all_income) & " - " & hours(all_income) & " hrs."
								If exclude_ALL_amount(all_income) = "" Then
									If exclude_from_SNAP(all_income) = unchecked Then
										snap_check_count = snap_check_count + 1
										snap_budgeted_total = snap_budgeted_total + gross_amount(all_income) - split_exclude_amount(all_income)
										snap_exclusion_total = split_exclude_amount(all_income)
										snap_hours = snap_hours + hours(all_income)
										If exclude_SNAP_amount(all_income) <> "" Then
											snap_budgeted_total = snap_budgeted_total - exclude_SNAP_amount(all_income)
											snap_exclusion_total = snap_exclusion_total + exclude_SNAP_amount(all_income)
											If IsNumeric(exclude_SNAP_hours(all_income)) Then snap_hours = snap_hours - (hourly_amount + exclude_SNAP_hours(all_income))
                                        Else
                                            exclude_SNAP_amount(all_income) = 0
                                        End If
										If snap_exclusion_total <> 0 Then SNAP_dialog_display(all_income) = SNAP_dialog_display(all_income)  & " - $ " & snap_exclusion_total & " not included."
									Else
										SNAP_dialog_display(all_income) = ""
										SNAP_list_of_excluded_pay_dates = SNAP_list_of_excluded_pay_dates & ", " & view_pay_date(all_income)    'making a list of all the checks that were not included in making the budget
									End If
									If exclude_from_CASH(all_income) = unchecked Then
										cash_check_count = cash_check_count + 1
										cash_budgeted_total = cash_budgeted_total + gross_amount(all_income) - split_exclude_amount(all_income)
										cash_exclusion_total = split_exclude_amount(all_income)
										cash_hours = cash_hours + hours(all_income)
										If exclude_CASH_amount(all_income) <> "" Then
											cash_budgeted_total = cash_budgeted_total - exclude_CASH_amount(all_income)
											cash_exclusion_total = cash_exclusion_total + exclude_CASH_amount(all_income)
											If IsNumeric(exclude_CASH_hours(all_income)) Then cash_hours = cash_hours - (hourly_amount + exclude_CASH_hours(all_income))
                                        Else
                                            exclude_CASH_amount(all_income) = 0
										End If
										If cash_exclusion_total <> 0 Then CASH_dialog_display(all_income) = CASH_dialog_display(all_income)  & " - $ " & cash_exclusion_total & " not included."
									Else
										CASH_dialog_display(all_income) = ""
										CASH_list_of_excluded_pay_dates = CASH_list_of_excluded_pay_dates & ", " & view_pay_date(all_income)    'making a list of all the checks that were not included in making the budget
									End If
								Else
									snap_check_count = snap_check_count + 1
									cash_check_count = cash_check_count + 1
									exclusion_total = exclude_ALL_amount(all_income) + split_exclude_amount(all_income)
									If exclusion_total <> 0 Then SNAP_dialog_display(all_income) = SNAP_dialog_display(all_income)  & " - $ " & exclusion_total & " not included."
									If exclusion_total <> 0 Then CASH_dialog_display(all_income) = CASH_dialog_display(all_income)  & " - $ " & exclusion_total & " not included."
									snap_budgeted_total = snap_budgeted_total + gross_amount(all_income) - exclusion_total
									cash_budgeted_total = cash_budgeted_total + gross_amount(all_income) - exclusion_total
									snap_hours = snap_hours + hours(all_income) - exclude_ALL_hours(all_income)
									cash_hours = cash_hours + hours(all_income) - exclude_ALL_hours(all_income)
								End If
							Else
								SNAP_list_of_excluded_pay_dates = SNAP_list_of_excluded_pay_dates & ", " & view_pay_date(all_income)    'making a list of all the checks that were not included in making the budget
								CASH_list_of_excluded_pay_dates = CASH_list_of_excluded_pay_dates & ", " & view_pay_date(all_income)    'making a list of all the checks that were not included in making the budget
							End If
						End If
					End If
				Next
			Next
			If SNAP_list_of_excluded_pay_dates <> "" Then SNAP_list_of_excluded_pay_dates = right(SNAP_list_of_excluded_pay_dates, len(SNAP_list_of_excluded_pay_dates) - 2)        'formatting this list to remove the leading ", "
			If CASH_list_of_excluded_pay_dates <> "" Then CASH_list_of_excluded_pay_dates = right(CASH_list_of_excluded_pay_dates, len(CASH_list_of_excluded_pay_dates) - 2)        'formatting this list to remove the leading ", "


			'This is a whole lot of math and formatting
			ave_hrs_per_pay = 0
			ave_inc_per_pay = 0
			cash_ave_hrs_per_pay = 0
			cash_ave_inc_per_pay = 0
			snap_ave_hrs_per_pay = 0
			snap_ave_inc_per_pay = 0
			If total_check_count <> 0 Then
				ave_hrs_per_pay = total_hours/total_check_count
                ave_inc_per_pay = ave_hrs_per_pay * hourly_wage
			End If
			If cash_check_count <> 0 Then
				cash_ave_hrs_per_pay = cash_hours/cash_check_count
				cash_ave_inc_per_pay = cash_budgeted_total/cash_check_count
			End If
			If snap_check_count <> 0 Then
				snap_ave_hrs_per_pay = snap_hours/snap_check_count
				snap_ave_inc_per_pay = snap_budgeted_total/snap_check_count
			End If
			ave_hrs_per_pay = FormatNumber(ave_hrs_per_pay, 2,,0)
			ave_inc_per_pay = FormatNumber(ave_inc_per_pay, 2,,0)
			cash_ave_hrs_per_pay = FormatNumber(cash_ave_hrs_per_pay, 2,,0)
			cash_ave_inc_per_pay = FormatNumber(cash_ave_inc_per_pay, 2,,0)
			snap_ave_hrs_per_pay = FormatNumber(snap_ave_hrs_per_pay, 2,,0)
			snap_ave_inc_per_pay = FormatNumber(snap_ave_inc_per_pay, 2,,0)

            If display_hrs_per_wk = "" Then
                snap_hrs_per_wk = 0
                'determining the number of hours per week for SNAP
                If pay_freq = "1 - One Time Per Month" Then snap_hrs_per_wk = snap_ave_hrs_per_pay/4.3
                If pay_freq = "2 - Two Times Per Month" Then snap_hrs_per_wk = (snap_ave_hrs_per_pay*2)/4.3
                If pay_freq = "3 - Every Other Week" Then snap_hrs_per_wk = snap_ave_hrs_per_pay/2
                If pay_freq = "4 - Every Week" Then snap_hrs_per_wk = snap_ave_hrs_per_pay

                cash_hrs_per_wk = 0
                'determining the number of hours per week for SNAP
                If pay_freq = "1 - One Time Per Month" Then cash_hrs_per_wk = cash_ave_hrs_per_pay/4.3
                If pay_freq = "2 - Two Times Per Month" Then cash_hrs_per_wk = (cash_ave_hrs_per_pay*2)/4.3
                If pay_freq = "3 - Every Other Week" Then cash_hrs_per_wk = cash_ave_hrs_per_pay/2
                If pay_freq = "4 - Every Week" Then cash_hrs_per_wk = cash_ave_hrs_per_pay

                hrs_per_wk = 0
                'determining the number of hours per week for non-SNAP
                If pay_freq = "1 - One Time Per Month" Then hrs_per_wk = ave_hrs_per_pay/4.3
                If pay_freq = "2 - Two Times Per Month" Then hrs_per_wk = (ave_hrs_per_pay*2)/4.3
                If pay_freq = "3 - Every Other Week" Then hrs_per_wk = ave_hrs_per_pay/2
                If pay_freq = "4 - Every Week" Then hrs_per_wk = ave_hrs_per_pay
            Else
                snap_hrs_per_wk = hrs_per_wk
                cash_hrs_per_wk = hrs_per_wk
            End If
            snap_hrs_per_wk = FormatNumber(snap_hrs_per_wk, 2,,0)
            cash_hrs_per_wk = FormatNumber(cash_hrs_per_wk, 2,,0)
            hrs_per_wk = FormatNumber(hrs_per_wk, 2,,0)

			monthly_income = ""
			CASH_monthly_income = ""
			SNAP_monthly_income = ""
			If pay_freq <> "" Then            'identifying the multiplier to determine monthly anticipated pay
				pay_multiplier = 0
				If pay_freq = "1 - One Time Per Month" Then pay_multiplier = 1
				If pay_freq = "2 - Two Times Per Month" Then pay_multiplier = 2
				If pay_freq = "3 - Every Other Week" Then pay_multiplier = 2.15
				If pay_freq = "4 - Every Week" Then pay_multiplier = 4.3
				monthly_income = pay_multiplier * ave_inc_per_pay     'monthly income
				CASH_monthly_income = pay_multiplier * cash_ave_inc_per_pay     'CASH monthly income
				SNAP_monthly_income = pay_multiplier * snap_ave_inc_per_pay     'SNAP monthly income

			End If

			If monthly_income = "" Then monthly_income = 0
			If CASH_monthly_income = "" Then CASH_monthly_income = 0
			If SNAP_monthly_income = "" Then SNAP_monthly_income = 0
			If monthly_income = 0 Then pick_one = use_estimate
			If CASH_monthly_income = 0 Then pick_one = use_estimate
			If SNAP_monthly_income = 0 Then pick_one = use_estimate
			monthly_income = FormatNumber(monthly_income, 2,,0)
			CASH_monthly_income = FormatNumber(CASH_monthly_income, 2,,0)
			SNAP_monthly_income = FormatNumber(SNAP_monthly_income, 2,,0)

		End If

		If pick_one = use_estimate Then
			'setting the wording for the CONFIRM BUDGET Dialog
			paycheck_list_title = "Anticipated Paychecks for " & initial_month_mo & "/" & initial_month_yr & ":"
			the_initial_month = DateValue(initial_month_mo & "/1/" & initial_month_yr)

			hourly_wage = pay_per_hr
			snap_hrs_per_wk = hrs_per_wk
			cash_hrs_per_wk = hrs_per_wk
			hourly_wage = FormatNumber(hourly_wage, 2,,0)
            snap_hourly_wage = hourly_wage
            cash_hourly_wage = hourly_wage

			snap_hrs_per_wk = FormatNumber(snap_hrs_per_wk, 2,,0)
			cash_hrs_per_wk = FormatNumber(cash_hrs_per_wk, 2,,0)

			ave_hrs_per_pay = 0
			ave_inc_per_pay = 0
			cash_ave_hrs_per_pay = 0
			cash_ave_inc_per_pay = 0
			snap_ave_hrs_per_pay = 0
			snap_ave_inc_per_pay = 0

			Select Case pay_freq          'here we determine averages of hours and income to anticipate based on pay frequency
				Case "1 - One Time Per Month"

					ave_inc_per_pay = pay_per_hr * hrs_per_wk * 4.3
					monthly_income = ave_inc_per_pay
					ave_hrs_per_pay = hrs_per_wk * 4.3
					default_start_date = the_initial_month
				Case "2 - Two Times Per Month"
                    If bimonthly_first = "" or bimonthly_second = "" then determine_bimonthly_dates
					ave_inc_per_pay = pay_per_hr * hrs_per_wk * 4.3 / 2
					monthly_income = ave_inc_per_pay * 2
					ave_hrs_per_pay = (hrs_per_wk * 4.3)/2
					default_start_date = DateValue(initial_month_mo & "/" & bimonthly_first & "/" & initial_month_yr)
				Case "3 - Every Other Week"
					ave_inc_per_pay = pay_per_hr * hrs_per_wk * 2
					monthly_income = ave_inc_per_pay * 2.15
					ave_hrs_per_pay = hrs_per_wk * 2
					the_date_of_week = the_initial_month
					Do
						If Weekday(the_date_of_week) = vbFriday Then
							default_start_date = the_date_of_week
							Exit Do
						Else
							the_date_of_week = DateAdd("d", 1, the_date_of_week)
						End If
					Loop
				Case "4 - Every Week"
					ave_inc_per_pay = pay_per_hr * hrs_per_wk
					monthly_income = ave_inc_per_pay * 4.3
					ave_hrs_per_pay = hrs_per_wk
					the_date_of_week = the_initial_month
					Do
						If Weekday(the_date_of_week) = vbFriday Then
							default_start_date = the_date_of_week
							Exit Do
						Else
							the_date_of_week =DateAdd("d", 1, the_date_of_week)
						End If
					Loop
			End Select
			ave_hrs_per_pay = FormatNumber(ave_hrs_per_pay, 2,,0)
			ave_inc_per_pay = FormatNumber(ave_inc_per_pay, 2,,0)
			monthly_income = FormatNumber(monthly_income, 2,,0)
			cash_ave_hrs_per_pay = ave_hrs_per_pay
			cash_ave_inc_per_pay = ave_inc_per_pay
			snap_ave_hrs_per_pay = ave_hrs_per_pay
			snap_ave_inc_per_pay = ave_inc_per_pay
			CASH_monthly_income = monthly_income
			SNAP_monthly_income = monthly_income

			' first_retro_check
			If first_check = "" Then
				If known_pay_date <> "" Then
					first_check = known_pay_date
				ElseIf actual_checks_provided Then
					first_check = first_check
				ElseIf first_retro_check <> "" Then
					first_check = first_retro_check
				Else
					first_check = default_start_date
				End If
			End If
		End If
		Call create_anticipated_pay_array
		' save_your_work

	end sub

	public sub case_note_details(developer_mode)
        'creates a CASE/NOTE for the job details

        If developer_mode = FALSE Then Call start_a_blank_CASE_NOTE        'now we start the case note

        If income_received Then        'if we have income verification - the note is more detailed
            STATS_manualtime = STATS_manualtime + 120
            If verif_type = "? - EXPEDITED SNAP ONLY" Then      'special header for if '?' is used as verification so they are easy to find
                Call write_variable_in_CASE_NOTE("XFS INCOME DETAIL: M" & member & " - JOBS - " & employer & " - PROG: " & prog_list)
            Else
                Call write_variable_in_CASE_NOTE("INCOME DETAIL: M" & member & " - JOBS - " & employer & " - PROG: " & prog_list)
            End If

            If new_panel = TRUE Then            'line in note about adding the panel
                Call write_variable_in_CASE_NOTE("* THIS IS NEW INCOME. Started on " & income_start_dt)
            End If
            Call write_variable_in_CASE_NOTE("*** - Pay Frequency: " & pay_freq & " - ***")

            If apply_to_SNAP = checked Then           'budget detail about SNAP
                If verif_type = "? - EXPEDITED SNAP ONLY"  Then Call write_variable_in_CASE_NOTE("BUDGET FOR SNAP ==================================================")
                If verif_type <> "? - EXPEDITED SNAP ONLY" Then Call write_variable_in_CASE_NOTE("BUDGET FOR SNAP ======================================= MONTHLY: $ " & SNAP_monthly_income)
                ' Call write_variable_in_CASE_NOTE("Anticipated Income Budget ----------------------------------")
                If verif_type = "? - EXPEDITED SNAP ONLY" Then          'different wording for if '?' verif code is used
                    Call write_variable_in_CASE_NOTE("*** JOBS has been updated with information that has not been verified. ***")
                    Call write_variable_in_CASE_NOTE("================================")
                    Call write_variable_in_CASE_NOTE("* Month of application: " & fs_appl_footer_month & "/" & fs_appl_footer_year & ". Income updated to determine eligibility for Expedited SNAP, which does not have to be verifed.")
                    Call write_variable_in_CASE_NOTE("-- Income for " & income_lumped_mo & " has been entered on PIC as a single monthly payment. --")
                    Call write_variable_in_CASE_NOTE("* Total income for this month - $" & lump_gross & " - " & lump_hrs & " hrs. Pay Frequency: Monthly")
                    Call write_variable_with_indent_in_CASE_NOTE("Income budgeted this way for this month because " & lump_reason)
                    Call write_variable_with_indent_in_CASE_NOTE("Amount determined using following checks:")
                    Call write_variable_with_indent_in_CASE_NOTE("-- Actual Checks: " & act_checks_lumped)
                    Call write_variable_with_indent_in_CASE_NOTE("-- Anticipated Checks: " & est_checks_lumped)
                Else
                    'There used to be separate handling for UHFS cases - but this should no longer be necessary
                    If SNAP_monthly_income = 0 Then Call write_variable_in_CASE_NOTE("!! THIS JOB IS NOT ANTICIPATING ANY INCOME AT THIS TIME. !!")
                    Call write_variable_in_CASE_NOTE("$ " & SNAP_monthly_income & " --- SNAP Monthly Budgeted Income.")
                    Call write_variable_in_CASE_NOTE("  *** Income has been reviewed and is anticipated to continue at this amount.")
                    Call write_variable_in_CASE_NOTE("  - Average per Pay Period: $" & snap_ave_inc_per_pay)
                    Call write_variable_in_CASE_NOTE("  - Average " & word_for_freq & " hours: " & snap_ave_hrs_per_pay)
                    Call write_variable_in_CASE_NOTE("  - Average pay per hour: $" & snap_hourly_wage & "/hr")
                    ' Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", pay_freq)

                    If income_lumped_mo <> "" Then
                        Call write_variable_in_CASE_NOTE("-- Income for " & income_lumped_mo & " has been entered on PIC as a single monthly payment. --")
                        Call write_variable_in_CASE_NOTE("* Total income for this month - $" & lump_gross & " - " & lump_hrs & " hrs. Pay Frequency: Monthly")
                        Call write_variable_in_CASE_NOTE("  - Income budgeted this way for this month because " & lump_reason)
                        Call write_variable_in_CASE_NOTE("  - Amount determined using following checks:")
                        Call write_variable_in_CASE_NOTE("    - Actual Checks: " & act_checks_lumped)
                        Call write_variable_in_CASE_NOTE("    - Anticipated Checks: " & est_checks_lumped)
                    End If
                End If
                'special notes about using the SNAP PIC
                If pick_one = use_actual Then Call write_variable_in_CASE_NOTE("* All included checks have been added to the SNAP PIC. Gross amount on PIC is reflective of the included pay amount.")
            End If

            If apply_to_CASH = checked or apply_to_GRH = checked Then           'Cash budget detail
                cash_header = ""
                If apply_to_CASH = checked and apply_to_GRH = checked Then            'if both SNAP and GRH are checked - we do a combined header
                    cash_header = "BUDGET FOR CASH and GRH ==============================="
                Else
                    If apply_to_CASH = checked Then cash_header = "BUDGET FOR CASH ======================================="
                    If apply_to_GRH = checked  Then cash_header = "BUDGET FOR GRH ========================================"
                End If
                If excl_cash_rsn = "" Then cash_header = cash_header & " MONTHLY: $ " & CASH_monthly_income
                If excl_cash_rsn <> "" Then cash_header = cash_header & "==========="
                Call write_variable_in_CASE_NOTE(cash_header)
                If excl_cash_rsn <> "" Then
                    Call write_variable_in_CASE_NOTE("* This income is not counted in the Cash budget. Reason: " & excl_cash_rsn)
                Else
                    If CASH_monthly_income = 0 Then Call write_variable_in_CASE_NOTE("!! THIS JOB IS NOT ANTICIPATING ANY INCOME AT THIS TIME. !!")
                    Call write_variable_in_CASE_NOTE("$ " & CASH_monthly_income & " --- CASH Monthly Budgeted Income.")
                    Call write_variable_in_CASE_NOTE("  - Average per Pay Period: $" & cash_ave_inc_per_pay)
                    Call write_variable_in_CASE_NOTE("  - Average " & word_for_freq & " hours: " & cash_ave_hrs_per_pay)
                    Call write_variable_in_CASE_NOTE("  - Average pay per hour: $" & cash_hourly_wage & "/hr")
                End If
                If pick_one = use_actual Then Call write_variable_in_CASE_NOTE("* All included checks have been added to the CASH PIC. Gross amount on PIC is reflective of the included pay amount.")
            End If

            If apply_to_HC = checked Then             'Health Care Budget Detail
                Call write_variable_in_CASE_NOTE("BUDGET FOR HEALTH CARE ================================")
                Call write_bullet_and_variable_in_CASE_NOTE("Average per Pay Period", "$" & ave_inc_per_pay)
                ' Call write_bullet_and_variable_in_CASE_NOTE("Pay Frequency", pay_freq)
                If hc_retro = TRUE Then
                    If cash_array_info_exists Then
                        For each_cash_month = 0 to UBOUND(cash_info_cash_mo_yr)
                            Call write_variable_in_CASE_NOTE("* Income updated in " & cash_info_cash_mo_yr(each_cash_month))
                            If cash_info_retro_updtd(each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("  -RETRO Income updated: $" & cash_info_mo_retro_pay(each_cash_month) & " total income for " & cash_info_retro_mo_yr(each_cash_month) & " with " & cash_info_mo_retro_hrs(each_cash_month) & " total hrs.")
                            If cash_info_prosp_updtd(each_cash_month) = TRUE Then Call write_variable_in_CASE_NOTE("  -Prosp Income updated: $" & cash_info_mo_prosp_pay(each_cash_month) & " total income for " & cash_info_cash_mo_yr(each_cash_month) & " with " & cash_info_mo_prosp_hrs(each_cash_month) & " total hrs.")
                        Next
                    End If
                End If
                Call write_bullet_and_variable_in_CASE_NOTE("Notes on HC Budget", hc_budg_notes)
            End If

            'Every program gets information about what verification was provided
            'Though most of this is really for SNAP requirements - it is still relevant to other programs
            Call write_variable_in_CASE_NOTE("VERIFICATION DETAILS === Received: " & verif_date & "================================")

            Call write_bullet_and_variable_in_CASE_NOTE("Received Date", verif_date)
            Call write_bullet_and_variable_in_CASE_NOTE("Type Received", verif_type)
            If verif_explain <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Explanation of Verification: " & verif_explain)

            Call write_bullet_and_variable_in_CASE_NOTE("Conversation with", spoke_with)
            If convo_detail <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Conversation Details: " & convo_detail)

            'Basically a list of the verification that was received
            Call write_variable_in_CASE_NOTE("INCOME INFORMATION RECEIVED ================================")

            If there_are_counted_checks AND anticipated_income_provided Then        'If there is an order ubound then there are actual checks'
                Call write_variable_in_CASE_NOTE("* Both actual check stubs and anticipated income estimates were received for this income.")

                If pick_one = use_actual Then Call write_variable_in_CASE_NOTE("* Actual pay amounts used to determine income to budget.")
                If pick_one = use_estimate Then Call write_variable_in_CASE_NOTE("* Income to budget determined by anticipated hours and rate of pay.")
                Call write_bullet_and_variable_in_CASE_NOTE("Reason for choice", selection_rsn)
            End If

            If there_are_counted_checks Then            'list of checks in order
                If verif_type = "? - EXPEDITED SNAP ONLY" Then
                    Call write_variable_in_CASE_NOTE("                            Pay information reported by client")
                Else
                    Call write_variable_in_CASE_NOTE("                            Checks provided to agency.")
                End If
                For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
                    For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
                        'conditional if it is the right panel AND the order matches - then do the thing you need to do
                        If check_order(all_income) = order_number Then
                            pay_info_string = ""
                            If exclude_entirely(all_income) Then
                                pay_info_string = pay_info_string & "** CHECK EXCLUDED FROM BUDGETS because " & reason_to_exclude(all_income) & "#%#"
                            Else
                                exclude_from_both = 0
                                exclude_from_both_rsn = ""
                                If IsNumeric(exclude_ALL_amount(all_income)) Then
                                    If exclude_ALL_amount(all_income) <> 0 Then
                                        exclude_from_both = exclude_from_both + exclude_ALL_amount(all_income)
                                        exclude_from_both_rsn = exclude_from_both_rsn & reason_to_exclude(all_income) & " "
                                    End If
                                End If
                                If IsNumeric(split_exclude_amount(all_income)) Then
                                    If split_exclude_amount(all_income) <> 0 Then
                                        exclude_from_both = exclude_from_both + split_exclude_amount(all_income)
                                        exclude_from_both_rsn = exclude_from_both_rsn & split_check_excld_string(all_income) & " "
                                    End If
                                End If

                                If apply_to_SNAP = checked and (apply_to_CASH = checked or apply_to_GRH = checked) Then
                                    sm_temp_string = ""
                                    If exclude_from_SNAP(all_income) = checked and exclude_from_CASH(all_income) = checked Then
                                        sm_temp_string = sm_temp_string & "** CHECK EXCLUDED FROM BUDGETS because " & reason_SNAP_amt_excluded(all_income) & " (SNAP) " & reason_CASH_amt_excluded(all_income) & "(CASH)"
                                    Else
                                        If exclude_from_SNAP(all_income) = checked Then sm_temp_string = sm_temp_string & "** THIS CHECK EXCLUDED FROM SNAP BUDGET because " & reason_SNAP_amt_excluded(all_income) & "#%#"
                                        If exclude_from_CASH(all_income) = checked Then sm_temp_string = sm_temp_string & "** THIS CHECK EXCLUDED FROM CASH BUDGET because " & reason_CASH_amt_excluded(all_income) & "#%#"
                                        sm_snap_excld_calc = 0
                                        sm_cash_excld_calc = 0
                                        sm_snap_excld_calc = sm_snap_excld_calc + exclude_SNAP_amount(all_income) + exclude_from_both
                                        sm_cash_excld_calc = sm_cash_excld_calc + exclude_CASH_amount(all_income) + exclude_from_both
                                        If sm_snap_excld_calc = sm_cash_excld_calc and sm_snap_excld_calc <> 0 Then
                                            sm_temp_string = sm_temp_string & "Only $" & gross_amount(all_income) - sm_snap_excld_calc & " included in SNAP and CASH budget because: " & exclude_from_both_rsn & " " & reason_SNAP_amt_excluded(all_income) & "#%#-- $" & sm_snap_excld_calc & " of check not included." & "#%#"
                                        Else
                                            If sm_snap_excld_calc <> 0 Then sm_temp_string = sm_temp_string & "Only $" & gross_amount(all_income) - sm_snap_excld_calc & " included in SNAP budget because: " & exclude_from_both_rsn & " " & reason_SNAP_amt_excluded(all_income) & "#%#-- $" & sm_snap_excld_calc & " of check not included." & "#%#"
                                            If sm_cash_excld_calc <> 0 Then sm_temp_string = sm_temp_string & "Only $" & gross_amount(all_income) - sm_cash_excld_calc & " included in CASH budget because: " & exclude_from_both_rsn & " " & reason_CASH_amt_excluded(all_income) & "#%#-- $" & sm_cash_excld_calc & " of check not included." & "#%#"
                                        End If
                                    End If
                                    pay_info_string = pay_info_string & sm_temp_string

                                ElseIf apply_to_SNAP = checked Then           'different formatting for different scenarios
                                    If exclude_from_SNAP(all_income) = checked Then
                                        pay_info_string = pay_info_string & "** CHECK EXCLUDED FROM SNAP BUDGET because " & reason_SNAP_amt_excluded(all_income) & "#%#"
                                    Else
                                        If NOT IsNumeric(exclude_SNAP_amount(all_income)) Then
                                            If exclude_SNAP_amount(all_income) = "" Then exclude_SNAP_amount(all_income) = 0
                                        End If
                                        If exclude_SNAP_amount(all_income) <> 0 or exclude_from_both <> 0 Then
                                            calc_snap_exclude = exclude_SNAP_amount(all_income) + exclude_SNAP_amount(all_income)
                                            calc_snap_exclude_rsn = exclude_from_both_rsn & " " & reason_SNAP_amt_excluded(all_income)
                                            pay_info_string = pay_info_string & "Only $" & gross_amount(all_income) - calc_snap_exclude & " included in SNAP budget because: " & trim(calc_snap_exclude_rsn) & "#%#-- $" & calc_snap_exclude & " of check not included." & "#%#"
                                        End If
                                    End If
                                ElseIf apply_to_CASH = checked or apply_to_GRH = checked Then           'different formatting for different scenarios
                                    If exclude_from_CASH(all_income) = checked Then
                                        pay_info_string = pay_info_string & "** CHECK EXCLUDED FROM CASH BUDGET because " & reason_CASH_amt_excluded(all_income) & "#%#"
                                    Else
                                        If NOT IsNumeric(exclude_CASH_amount(all_income)) Then
                                            If exclude_CASH_amount(all_income) = "" Then exclude_CASH_amount(all_income) = 0
                                        End If
                                        If exclude_CASH_amount(all_income) <> 0 or exclude_from_both <> 0 Then
                                            calc_cash_exclude = exclude_CASH_amount(all_income) + exclude_from_both
                                            calc_cash_exclude_rsn = exclude_from_both_rsn & " " & reason_CASH_amt_excluded(all_income)
                                            pay_info_string = pay_info_string & "Only $" & gross_amount(all_income) - calc_cash_exclude & " included in CASH budget because: " & trim(calc_cash_exclude_rsn) & "#%#-- $" & calc_cash_exclude & " of check not included." & "#%#"
                                        End If
                                    End If
                                End If
                            End If
                            Call write_variable_in_CASE_NOTE("Paydate: " & right("      "&view_pay_date(all_income), 10) & " --- Gross: $ " & right("          "&gross_amount(all_income), gross_max_string_len) & " --- Hours Worked: " & hours(all_income))
                            pay_lines = split(pay_info_string, "#%#")
                            For each pay_info in pay_lines
                                If left(pay_info, 2) = "**" Then
                                    Call write_variable_in_CASE_NOTE("  " & pay_info)
                                Else
                                    Call write_variable_with_indent_in_CASE_NOTE(pay_info)
                                End If
                            Next

                            If combined_into_one(all_income) Then Call write_variable_in_CASE_NOTE("  - This check was combined with all from the same date on JOBS.")
                            If bonus_check(all_income)       Then Call write_variable_in_CASE_NOTE("  - This is a BONUS CHECK.")
                            If pay_detail_exists(all_income) = True Then
                                Call write_variable_in_CASE_NOTE("  - The pay information for this check is split:")
                                Call write_variable_in_CASE_NOTE("    *Regular Pay: $ " & pay_split_regular_amount(all_income))
                                If pay_split_bonus_amount(all_income) <> "" Then Call write_variable_in_CASE_NOTE("    *Bonus Pay: $ " & pay_split_bonus_amount(all_income))
                                If pay_split_ot_amount(all_income) <> "" Then Call write_variable_in_CASE_NOTE("    *OT Pay: $ " & pay_split_ot_amount(all_income))
                                If pay_split_shift_diff_amount(all_income) <> "" Then Call write_variable_in_CASE_NOTE("    *Shift Differential Pay: $ " & pay_split_shift_diff_amount(all_income))
                                If pay_split_tips_amount(all_income) <> "" Then Call write_variable_in_CASE_NOTE("    *Tip Pay: $ " & pay_split_tips_amount(all_income))
                                If pay_split_other_amount(all_income) <> "" Then Call write_variable_in_CASE_NOTE("    *" & pay_split_other_detail(all_income) &" Pay: $ " & pay_split_other_amount(all_income))
                                'The excluded detail of the pay split should be in the xcluded string above and does not need to be handled here
                            End If
                            If calculated_by_ytd(all_income) then
                                Call write_variable_in_CASE_NOTE(" - Calculated using YTDs")
                                note_line_array = split(ytd_calc_notes(all_income), ";")
                                for each ytd_info_line in note_line_array
                                    Call write_variable_in_CASE_NOTE("      " & trim(ytd_info_line))
                                next
                            End if
                            If future_check(all_income) Then Call write_variable_in_CASE_NOTE("        Pay Date in future - reported expected amount, Only used for SNAP budget in month of application.")

                        End If
                    next
                next
            End If

            mo_w_more_5_chcks = trim(mo_w_more_5_chcks)
            If mo_w_more_5_chcks <> "" Then
                Call write_variable_in_CASE_NOTE("* These months have more than 5 paychecks and were entered into JOBS")
                Call write_variable_in_CASE_NOTE("  as a single amount: " & mo_w_more_5_chcks)
            End If
            If anticipated_income_provided Then
                Call write_variable_in_CASE_NOTE("* Anticipated Income Estimate provided to Agency.")

                Call write_variable_with_indent_in_CASE_NOTE("Hourly Pay Rate: $" & pay_per_hr & "/hr")
                Call write_variable_with_indent_in_CASE_NOTE("Hours Per Week: " & hrs_per_wk & " hours")
                Call write_variable_with_indent_in_CASE_NOTE("Pay Frequency: " & pay_freq)
            End If

            Call write_variable_in_CASE_NOTE("ACTION TAKEN: JOBS Updated ================================")         'currently very basic
            If left(months_updated, 1) = "," Then  months_updated = right(months_updated, len(months_updated) - 1)

            Call write_bullet_and_variable_in_CASE_NOTE("Months updated", months_updated)
            If apply_to_CASH = checked  OR hc_retro = TRUE Then Call write_variable_in_CASE_NOTE("* Actual Checks entered on RETRO Side of the MAIN panel if received.")
        Else        'this is if there was ONLY the panel added but no verification of income was entered
            Call write_variable_in_CASE_NOTE("New Job: M" & member & " - JOBS - " & employer & " - PROG: " & prog_list)
            Call write_variable_in_CASE_NOTE("* Information received that a new job has started.")
        End If
        Call write_variable_in_CASE_NOTE("---")
        Call write_variable_in_CASE_NOTE(worker_signature)

	end sub

	public sub check_details_dialog(edit_base_info, check_count, ButtonPressed)
        'dialog to record more details about a check like split in pay details or exclusions

        Do

            If edit_base_info = True Then
                pay_date(check_count) = pay_date(check_count) & ""
                gross_amount(check_count) = gross_amount(check_count) & ""
                hours(check_count) = hours(check_count) & ""
            End If
            bonus_exclude_txt = "Bonus Check is not expected to continue."
            pay_split_regular_amount(check_count) = pay_split_regular_amount(check_count) & ""
            pay_split_bonus_amount(check_count) = pay_split_bonus_amount(check_count) & ""
            pay_split_ot_amount(check_count) = pay_split_ot_amount(check_count) & ""
            pay_split_shift_diff_amount(check_count) = pay_split_shift_diff_amount(check_count) & ""
            pay_split_tips_amount(check_count) = pay_split_tips_amount(check_count) & ""
            pay_split_other_amount(check_count) = pay_split_other_amount(check_count) & ""

            exclude_ALL_amount(check_count) = exclude_ALL_amount(check_count) & ""
            exclude_SNAP_amount(check_count) = exclude_SNAP_amount(check_count) & ""
            exclude_CASH_amount(check_count) = exclude_CASH_amount(check_count) & ""

            exclude_ALL_hours(check_count) = exclude_ALL_hours(check_count) & ""
            exclude_SNAP_hours(check_count) = exclude_SNAP_hours(check_count) & ""
            exclude_CASH_hours(check_count) = exclude_CASH_hours(check_count) & ""

            If exclude_ALL_amount(check_count) = "0" Then exclude_ALL_amount(check_count) = ""
            If exclude_SNAP_amount(check_count) = "0" Then exclude_SNAP_amount(check_count) = ""
            If exclude_CASH_amount(check_count) = "0" Then exclude_CASH_amount(check_count) = ""

            If bonus_check(check_count) = True Then bonus_checkbox = checked
            If exclude_entirely(check_count) = True Then exclude_all_checkbox = checked

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 570, 280, "Check Details"
                If edit_base_info = True Then
                    GroupBox 10, 5, 220, 65, "Actual Check Details"
                    Text 20, 20, 35, 10, "Pay Date:"
                    EditBox 55, 15, 50, 15, pay_date(check_count)
                    Text 115, 20, 60, 10, "Gross Amount: $"
                    EditBox 170, 15, 50, 15, gross_amount(check_count)
                    Text 125, 40, 45, 10, "Total Hours:"
                    EditBox 170, 35, 50, 15, hours(check_count)
                    Text 55, 30, 50, 10, "(MM/DD/YY)"
                    CheckBox 55, 55, 150, 10, "Check Here if this is a BONUS CHECK", bonus_checkbox
                    ' GroupBox 10, 70, 225, 170, "Gross Pay Split - OPTIONAL"
                End If
                If edit_base_info = False Then
                    GroupBox 10, 10, 115, 50, "Actual Check Details"
                    Text 30, 25, 90, 10, "Pay Date: " & pay_date(check_count)
                    Text 15, 35, 85, 10, "Gross Amount: $ " & gross_amount(check_count)
                    Text 25, 45, 85, 10, "Total Hours: " & hours(check_count)
                    If bonus_check(check_count) = True Then Text 130, 25, 90, 10 , "* THIS IS A BONUS CHECK *"
                End If
                CheckBox 245, 10, 235, 10, "Check Here to Exclude the entire check for SNAP, CASH, and GRH", exclude_all_checkbox
                Text 245, 30, 55, 10, "Exclude Portion:"
                EditBox 300, 25, 50, 15, exclude_ALL_amount(check_count)
                Text 350, 30, 105, 10, "Exclude portion of hours as well"
                EditBox 455, 25, 40, 15, exclude_ALL_hours(check_count)
                Text 245, 45, 90, 10, "Explain Exclusion Reason:"
                EditBox 245, 55, 315, 15, reason_to_exclude(check_count)

                GroupBox 10, 80, 225, 175, "Gross Pay Split - OPTIONAL"
                EditBox 65, 95, 50, 15, pay_split_regular_amount(check_count)
                EditBox 65, 125, 50, 15, pay_split_bonus_amount(check_count)
                CheckBox 130, 130, 95, 10, "Exclude Bonus Portion", pay_excld_bonus(check_count)
                EditBox 65, 145, 35, 15, pay_split_ot_amount(check_count)
                EditBox 125, 145, 25, 15, pay_split_ot_hours(check_count)
                CheckBox 155, 150, 95, 10, "Exclude OT", pay_excld_ot(check_count)
                EditBox 65, 165, 50, 15, pay_split_shift_diff_amount(check_count)
                CheckBox 130, 170, 95, 10, "Exclude Shift Diff. Portion", pay_excld_shift_diff(check_count)
                EditBox 65, 185, 50, 15, pay_split_tips_amount(check_count)
                CheckBox 130, 190, 95, 10, "Exclude Tips Portion", pay_excld_tips(check_count)
                EditBox 65, 205, 50, 15, pay_split_other_amount(check_count)
                CheckBox 130, 210, 90, 10, "Exclude Other Portion", pay_excld_other(check_count)
                EditBox 130, 225, 95, 15, pay_split_other_detail(check_count)
                Text 15, 100, 50, 10, "Regular Pay: $"
                Text 15, 110, 205, 10, "** Exclusions in a pay split apply to BOTH SNAP and CASH"
                Text 20, 130, 45, 10, "Bonus Pay:  $"
                Text 30, 150, 35, 10, "OT Pay:  $"
                Text 105, 150, 20, 10, "hours"
                Text 10, 170, 55, 10, "Shift Diff. Pay:  $"
                Text 25, 190, 40, 10, "Tips Pay:  $"
                Text 20, 210, 45, 10, "Other Pay:  $"
                Text 75, 230, 55, 10, "Explain 'Other':"
                CheckBox 270, 90, 200, 10, "Check here to exclude entire check for SNAP only.", exclude_from_SNAP(check_count)
                EditBox 345, 115, 50, 15, exclude_SNAP_amount(check_count)
                EditBox 505, 115, 50, 15, exclude_SNAP_hours(check_count)
                EditBox 255, 145, 300, 15, reason_SNAP_amt_excluded(check_count)
                GroupBox 245, 80, 315, 85, "SNAP"
                GroupBox 245, 105, 315, 60, "SNAP Only Exclusions"
                Text 255, 120, 90, 10, "SNAP Exclusion Amount: $"
                Text 400, 120, 105, 10, "Exclude portion of hours as well"
                Text 255, 135, 85, 10, "SNAP Exclusion Reason:"
                CheckBox 270, 180, 200, 10, "Check here to exclude entire check for CASH only.", exclude_from_CASH(check_count)
                EditBox 345, 205, 50, 15, exclude_CASH_amount(check_count)
                EditBox 505, 205, 50, 15, exclude_CASH_hours(check_count)
                EditBox 255, 235, 300, 15, reason_CASH_amt_excluded(check_count)
                GroupBox 245, 170, 315, 85, "CASH"
                GroupBox 245, 195, 315, 60, "CASH Only Exclusions"
                Text 255, 210, 90, 10, "CASH Exclusion Amount: $"
                Text 400, 210, 105, 10, "Exclude portion of hours as well"
                Text 255, 225, 85, 10, "CASH Exclusion Reason:"
                ButtonGroup ButtonPressed
                    PushButton 465, 260, 95, 15, "Keep Check Details", save_details_btn
                    PushButton 10, 260, 75, 15, "Delete Check", delete_check_btn
            EndDialog

			err_msg = ""
			dialog Dialog1
            save_your_work

			If ButtonPressed = -1 Then ButtonPressed = save_details_btn
			If ButtonPressed = 0 Then ButtonPressed = delete_check_btn

			If pay_split_regular_amount(check_count) = "" and pay_split_bonus_amount(check_count) = "" and pay_split_ot_amount(check_count) = "" and pay_split_shift_diff_amount(check_count) = "" and pay_split_tips_amount(check_count) = "" and pay_split_other_amount(check_count) = "" Then
				pay_split_regular_amount(check_count) = gross_amount(check_count)
			End If

			If edit_base_info = True Then
				bonus_check(check_count) = False
				If bonus_checkbox = checked then
                    bonus_check(check_count) = True
                    If exclude_all_checkbox = checked and trim(reason_to_exclude(check_count)) = "" Then reason_to_exclude(check_count) = "Bonus Check is not expected to continue."
                End If
			End If

            exclude_entirely(check_count) = False
            If exclude_all_checkbox = checked then exclude_entirely(check_count) = True

			Call check_details_error_handling(err_msg, edit_base_info)

			If ButtonPressed = delete_check_btn Then err_msg = ""

			If err_msg <> "" Then MsgBox "*  *  *  NOTICE  *  *  *" & vbCr & "Please Update the information in the dialog to continue:" & vbCr & err_msg

		Loop until err_msg = ""

		If edit_base_info = True and ButtonPressed <> delete_check_btn Then
			pay_date(check_count) = DateAdd("d", 0, pay_date(check_count))
			gross_amount(check_count) = FormatNumber(gross_amount(check_count), 2, -1, 0, 0)
			bonus_check(check_count) = False
			If bonus_checkbox = checked then bonus_check(check_count) = True
		End If

		pay_detail_exists(check_count) = False
		' If IsNumeric(pay_split_regular_amount(check_count)) = True Then pay_detail_exists(check_count) = True
		If IsNumeric(pay_split_bonus_amount(check_count)) = True Then pay_detail_exists(check_count) = True
		If IsNumeric(pay_split_ot_amount(check_count)) = True Then pay_detail_exists(check_count) = True
		If IsNumeric(pay_split_shift_diff_amount(check_count)) = True Then pay_detail_exists(check_count) = True
		If IsNumeric(pay_split_tips_amount(check_count)) = True Then pay_detail_exists(check_count) = True
		If IsNumeric(pay_split_other_amount(check_count)) = True Then pay_detail_exists(check_count) = True
        If pay_detail_exists(check_count) and NOT IsNumeric(pay_split_regular_amount(check_count)) Then
            split_amount = 0
            If IsNumeric(pay_split_bonus_amount(check_count))       Then split_amount = split_amount + pay_split_bonus_amount(check_count)*1
            If IsNumeric(pay_split_ot_amount(check_count))          Then split_amount = split_amount + pay_split_ot_amount(check_count)*1
            If IsNumeric(pay_split_shift_diff_amount(check_count))  Then split_amount = split_amount + pay_split_shift_diff_amount(check_count)*1
            If IsNumeric(pay_split_tips_amount(check_count))        Then split_amount = split_amount + pay_split_tips_amount(check_count)*1
            If IsNumeric(pay_split_other_amount(check_count))       Then split_amount = split_amount + pay_split_other_amount(check_count)*1
            pay_split_regular_amount(check_count) = gross_amount(check_count)*1 - split_amount
			pay_split_regular_amount(check_count) = FormatNumber(pay_split_regular_amount(check_count), 2, -1, 0, 0)
        End If
	end sub

	public sub check_details_error_handling(err_msg, edit_base_info)
        'making sure the details of check align with data requirements and formats

		If edit_base_info = True Then
			If NOT IsDate(pay_date(check_count)) Then err_msg = err_msg & vbCr & "* Pay Date should be entered as a date."
			If IsDate(pay_date(check_count)) Then
				future_check(check_count) = False
                in_appl_month = False
                If IsDate(fs_appl_date) Then
                    If DatePart("m", fs_appl_date) = DatePart("m", pay_date(check_count)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", pay_date(check_count)) Then   'if the paydate is in the application month
                        in_appl_month = True
                    End If
                End If
				If in_appl_month Then   'if the paydate is in the application month
					If DateDiff("d", date, pay_date(check_count)) > 0 Then future_check(check_count) = TRUE   'this is a future check
				Else        'if the paydate is NOT in the application  month
					If DateDiff("d", date, pay_date(check_count)) > 0 Then             'if the pay date is in the future we have to error
						err_msg = err_msg & vbCr & "* Paydates cannot be in the future. (" & pay_date(check_count) & ")"
					End If
				End If
			End If
			If NOT IsNumeric(gross_amount(check_count)) Then err_msg = err_msg & vbCr & "* The Amount of the Pay Check should be entered."
			If NOT IsNumeric(hours(check_count)) Then err_msg = err_msg & vbCr & "* The Hours worked should be entered as a number."
		End If
		If trim(exclude_ALL_amount(check_count)) <> "" and NOT IsNumeric(exclude_ALL_amount(check_count)) Then err_msg = err_msg & vbCr & "* The Exclude ALL Amount has been entered but does not appear to be a number, this must be entered as a number."
		If trim(exclude_ALL_hours(check_count)) <> "" and NOT IsNumeric(exclude_ALL_hours(check_count)) Then err_msg = err_msg & vbCr & "* The Exclude ALL Hours has been entered but does not appear to be a number, this must be entered as a number."
		If trim(exclude_SNAP_amount(check_count)) <> "" and NOT IsNumeric(exclude_SNAP_amount(check_count)) Then err_msg = err_msg & vbCr & "* The Exclude SNAP Amount has been entered but does not appear to be a number, this must be entered as a number."
		If trim(exclude_SNAP_hours(check_count)) <> "" and NOT IsNumeric(exclude_SNAP_hours(check_count)) Then err_msg = err_msg & vbCr & "* The Exclude SNAP Hours has been entered but does not appear to be a number, this must be entered as a number."
		If trim(exclude_CASH_amount(check_count)) <> "" and NOT IsNumeric(exclude_CASH_amount(check_count)) Then err_msg = err_msg & vbCr & "* The Exclude CASH Amount has been entered but does not appear to be a number, this must be entered as a number."
		If trim(exclude_SNAP_hours(check_count)) <> "" and NOT IsNumeric(exclude_SNAP_hours(check_count)) Then err_msg = err_msg & vbCr & "* The Exclude CASH Hours has been entered but does not appear to be a number, this must be entered as a number."

        If exclude_ALL_hours(check_count) = "" Then exclude_ALL_hours(check_count) = 0
        If IsNumeric(exclude_SNAP_amount(check_count)) Then
            If exclude_SNAP_amount(check_count) = 0 Then exclude_SNAP_amount(check_count) = ""
        End If
        If IsNumeric(exclude_CASH_amount(check_count)) Then
            If exclude_CASH_amount(check_count) = 0 Then exclude_CASH_amount(check_count) = ""
        End If

		If exclude_from_SNAP(check_count) = checked and exclude_from_CASH(check_count) = checked then
			exclude_all_checkbox = checked
			If trim(reason_to_exclude(check_count)) = "" Then
				If trim(reason_CASH_amt_excluded(check_count)) <> "" Then reason_to_exclude(check_count) = reason_CASH_amt_excluded(check_count)
				If trim(reason_SNAP_amt_excluded(check_count)) <> "" Then reason_to_exclude(check_count) = reason_SNAP_amt_excluded(check_count)
			End If
		End If
		If bonus_check(check_count) = True and exclude_all_checkbox = checked Then
			If Instr(reason_to_exclude(check_count), bonus_exclude_txt) = 0 Then
				reason_to_exclude(check_count) = bonus_exclude_txt
			End If
		End If
		If exclude_all_checkbox = checked or IsNumeric(exclude_ALL_amount(check_count)) Then
			'need to explain excluding a check
			If trim(reason_to_exclude(check_count)) = "" Then err_msg = err_msg & vbCr & "* To exclude the entire check, list a reason for excluding it."
		Else
			If (exclude_from_SNAP(check_count) = checked or (IsNumeric(exclude_SNAP_amount(check_count)))) and trim(reason_SNAP_amt_excluded(check_count)) = "" Then err_msg = err_msg & vbCr & "* Explain the reason for excluding part or all of the check for SNAP."
			If (exclude_from_CASH(check_count) = checked or (IsNumeric(exclude_CASH_amount(check_count)))) and trim(reason_CASH_amt_excluded(check_count)) = "" Then err_msg = err_msg & vbCr & "* Explain the reason for excluding part or all of the check for CASH."
		End If


		total_pay_calculation = 0

		pay_split_regular_amount(check_count) = trim(pay_split_regular_amount(check_count))
		If IsNumeric(pay_split_regular_amount(check_count)) = True Then
			pay_split_regular_amount(check_count) = pay_split_regular_amount(check_count)*1
			total_pay_calculation = total_pay_calculation + pay_split_regular_amount(check_count)
		Else
			If pay_split_regular_amount(check_count) <> "" Then err_msg = err_msg & vbCr & "* REGULAR Pay was entered but does not appear to be a valid number, please review."
		End If

		pay_split_bonus_amount(check_count) = trim(pay_split_bonus_amount(check_count))
		pay_split_ot_amount(check_count) = trim(pay_split_ot_amount(check_count))
		pay_split_shift_diff_amount(check_count) = trim(pay_split_shift_diff_amount(check_count))
		pay_split_tips_amount(check_count) = trim(pay_split_tips_amount(check_count))
		pay_split_other_amount(check_count) = trim(pay_split_other_amount(check_count))

		If IsNumeric(pay_split_bonus_amount(check_count)) = True Then
			pay_split_bonus_amount(check_count) = pay_split_bonus_amount(check_count)*1
			total_pay_calculation = total_pay_calculation + pay_split_bonus_amount(check_count)
		Else
			If pay_split_bonus_amount(check_count) <> "" Then err_msg = err_msg & vbCr & "* BONUS Pay was entered but does not appear to be a valid number, please review."
			If pay_excld_bonus(check_count) = checked Then err_msg = err_msg & vbCr & "* Exclude BONUS Pay was checked but amount entered does not appear to be a number."
		End If
		If IsNumeric(pay_split_ot_amount(check_count)) = True Then
			pay_split_ot_amount(check_count) = pay_split_ot_amount(check_count)*1
			total_pay_calculation = total_pay_calculation + pay_split_ot_amount(check_count)
		Else
			If pay_split_ot_amount(check_count) <> "" Then err_msg = err_msg & vbCr & "* OVERTIME Pay was entered but does not appear to be a valid number, please review."
			If pay_excld_ot(check_count) = checked Then err_msg = err_msg & vbCr & "* Exclude OVERTIME Pay was checked but amount entered does not appear to be a number."
		End If
		If IsNumeric(pay_split_shift_diff_amount(check_count)) = True Then
			pay_split_shift_diff_amount(check_count) = pay_split_shift_diff_amount(check_count)*1
			total_pay_calculation = total_pay_calculation + pay_split_shift_diff_amount(check_count)
		Else
			If pay_split_shift_diff_amount(check_count) <> "" Then err_msg = err_msg & vbCr & "* SHIFT DIFFERENTIAL Pay was entered but does not appear to be a valid number, please review."
			If pay_excld_shift_diff(check_count) = checked Then err_msg = err_msg & vbCr & "* Exclude SHIFT DIFFERENTIAL Pay was checked but amount entered does not appear to be a number."
		End If
		If IsNumeric(pay_split_tips_amount(check_count)) = True Then
			pay_split_tips_amount(check_count) = pay_split_tips_amount(check_count)*1
			total_pay_calculation = total_pay_calculation + pay_split_tips_amount(check_count)
		Else
			If pay_split_tips_amount(check_count) <> "" Then err_msg = err_msg & vbCr & "* TIPS Pay was entered but does not appear to be a valid number, please review."
			If pay_excld_tips(check_count) = checked Then err_msg = err_msg & vbCr & "* Exclude TIPS Pay was checked but amount entered does not appear to be a number."
		End If
		If IsNumeric(pay_split_other_amount(check_count)) = True Then
			pay_split_other_amount(check_count) = pay_split_other_amount(check_count)*1
			total_pay_calculation = total_pay_calculation + pay_split_other_amount(check_count)
			If trim(pay_split_other_detail(check_count)) = "" Then err_msg = err_msg & vbCr & "An amount was listed in OTHER Pay but no detail was entered into the explanation of what OTHER is. Update the explanation."
		Else
			If pay_split_other_amount(check_count) <> "" Then err_msg = err_msg & vbCr & "* OTHER (" & pay_split_other_detail(check_count) & ") Pay was entered but does not appear to be a valid number, please review."
			If pay_excld_other(check_count) = checked Then err_msg = err_msg & vbCr & "* Exclude OTHER (" & pay_split_other_detail(check_count) & ") Pay was checked but amount entered does not appear to be a number."
		End If
		total_pay_calculation = FormatNumber(total_pay_calculation, 2, -1, 0, 0)
		If IsNumeric(gross_amount(check_count)) Then
			gross_amount(check_count) = FormatNumber(gross_amount(check_count), 2, -1, 0, 0)
            If pay_split_regular_amount(check_count) = "" Then
                pay_split_regular_amount(check_count) = gross_amount(check_count) - total_pay_calculation
                total_pay_calculation = total_pay_calculation + pay_split_regular_amount(check_count)
        		total_pay_calculation = FormatNumber(total_pay_calculation, 2, -1, 0, 0)
            End If
			If total_pay_calculation <> gross_amount(check_count) Then
				err_msg = err_msg & vbCr & "* The pay entered in the split pay information does not match the gross pay amount entered. Update the numbers on the pay splits, or press the 'Clear' button to cancel the split pay functionality and return to the main Paycheck Received dialog to update the Gross Pay amount."
				err_msg = err_msg & vbCr & " - Gross Paycheck Amount $ " & gross_amount(check_count)
				err_msg = err_msg & vbCr & " - Sum of Paycheck Splits $ " & total_pay_calculation
			End If
		End If

	end sub

	public sub create_anticipated_pay_array()
        'this is to create a list of anticipated paychecks for the SNAP and CASH budget display

        days_to_add = 0     'this is for counting one check to the next
		months_to_add = 0

		Select Case pay_freq          'here we determine averages of hours and income to anticipate based on pay frequency
			Case "1 - One Time Per Month"
				days_to_add = 0
				months_to_add = 1
			Case "2 - Two Times Per Month"
				days_to_add = 30
				months_to_add = 1
			Case "3 - Every Other Week"
				days_to_add = 14
				months_to_add = 0
			Case "4 - Every Week"
				days_to_add = 7
				months_to_add = 0
		End Select

		the_initial_month = DateValue(initial_month_mo & "/1/" & initial_month_yr)
		snap_anticipated_pay_array = ""     'blanking these out because of looping
		cash_anticipated_pay_array = ""
		snap_checks_list = ""
		cash_checks_list = ""
		save_dates = FALSE
		this_pay_date = default_start_date
		this_pay_date = DateAdd("d", 0 , this_pay_date)
		Do      'While DatePart("m", this_pay_date) <> CM_2_mo AND DatePart("yyyy", this_pay_date) <> CM_2_yr
			save_dates = FALSE

			'if the date we are looking at is for the initial month - then we are going to save it to a list.
			If DatePart("m", this_pay_date) = DatePart("m", the_initial_month) AND DatePart("yyyy", this_pay_date) = DatePart("yyyy", the_initial_month) Then save_dates = TRUE
			If save_dates = TRUE Then

				check_found = FALSE         'looking to see if there was an actual check for this date
				For all_income = 0 to UBound(pay_date)
					If pay_date(all_income) <> "" Then
						If DateValue(pay_date(all_income)) = this_pay_date Then
							check_found = TRUE          'if there was then we will save that information for our list
							check_number = all_income   'need to know what position it is at
							Exit For
						End If
					End If
				Next
				'BUGGY CODE - we may need better handling for creating a list for the non-snap program sidplays
				If check_found = TRUE Then  'if the check was listed the information will include the actual amount listed
					'this different formatting is just to make it pretty
					If len(this_pay_date) = 10 Then checks_list = checks_list & "%" & this_pay_date & "   ~   $" & gross_amount(check_number)
					If len(this_pay_date) = 9 Then checks_list = checks_list & "%" & this_pay_date & "    ~   $" & gross_amount(check_number)
					If len(this_pay_date) = 8 Then checks_list = checks_list & "%" & this_pay_date & "     ~   $" & gross_amount(check_number)
				Else            'otherwise it includes a paycheck average
					If len(this_pay_date) = 10 Then checks_list = checks_list & "%" & this_pay_date & "   ~   $" & snap_ave_inc_per_pay
					If len(this_pay_date) = 9 Then checks_list = checks_list & "%" & this_pay_date & "    ~   $" & snap_ave_inc_per_pay
					If len(this_pay_date) = 8 Then checks_list = checks_list & "%" & this_pay_date & "     ~   $" & snap_ave_inc_per_pay
				End If
			End If
			If months_to_add = 0 Then       'these are defined by the pay frequency above and will increment us to the next pay date
				this_pay_date = DateAdd("d", days_to_add, this_pay_date)
			ElseIf days_to_add = 0 Then
				this_pay_date = DateAdd("m", months_to_add, this_pay_date)
			Else
				If bimonthly_second = "LAST" Then
					month_ahead = DateAdd("m", 1, this_pay_date)
					month_to_use = DatePart("m", month_ahead)
					year_to_use = DatePart("yyyy", month_ahead)
					If DatePart("d", this_pay_date) = bimonthly_first Then
						first_of_nextMonth = month_to_use & "/1/" & year_to_use
						this_pay_date = DateAdd("d", -1, first_of_nextMonth)
					Else
						this_pay_date = month_to_use & "/" & bimonthly_first & "/" & year_to_use
					End If
				Else
					If DatePart("d", this_pay_date) = bimonthly_first Then
						month_to_use = DatePart("m", this_pay_date)
						year_to_use = DatePart("yyyy", this_pay_date)
						this_pay_date = month_to_use & "/" & bimonthly_second & "/" & year_to_use
					ElseIf DatePart("d", this_pay_date) = bimonthly_second Then
						month_ahead = DateAdd("m", 1, this_pay_date)
						month_to_use = DatePart("m", month_ahead)
						year_to_use = DatePart("yyyy", month_ahead)
						this_pay_date = month_to_use & "/" & bimonthly_first & "/" & year_to_use
					End If
				End If
			End If
		Loop until DatePart("m", this_pay_date) = CM_2_mo AND DatePart("yyyy", this_pay_date) = CM_2_yr     'stop at current month plus 2

		'Formatting the list and making it an array
		If left(checks_list, 1) = "%" Then checks_list = right(checks_list, len(checks_list)-1)
		If InStr(checks_list, "%") <> 0 Then
			snap_anticipated_pay_array = Split(checks_list,"%")			'this is the array that is used in the confirmation dialog of the anticipated pay dates
			cash_anticipated_pay_array = Split(checks_list,"%")
		Else
			snap_anticipated_pay_array = Array(checks_list)
			cash_anticipated_pay_array = Array(checks_list)
		End If

	end sub

	public sub create_expected_check_array()
        'creates a list of pay dates that are expected for the job - needed to ensure all checks are provided

        call determine_bimonthly_dates
		call make_known_date_earlier

		list_of_all_paydates_start_to_finish = ""   'Here we loop through to create a list of all the paychcks that we should see from the first listed to the last
		next_paydate = first_check

		counter = 0
		Do
			list_of_all_paydates_start_to_finish = list_of_all_paydates_start_to_finish & "~" & next_paydate

			If pay_freq = "1 - One Time Per Month" Then       'each next date is determined by the pay frequency
				next_paydate = DateAdd("m", 1, next_paydate)
			ElseIf pay_freq = "2 - Two Times Per Month" Then
				If DatePart("d", next_paydate) = bimonthly_first Then         'If we are at the first check of the month, we need to go to the second
					next_pay_month = DatePart("m", next_paydate)
					next_pay_year = DatePart("yyyy", next_paydate)

					If bimonthly_second = "LAST" Then
						month_after = next_pay_month & "/1/" & next_pay_year
						month_after = DateAdd("m", 1, month_after)
						next_paydate = DateAdd("d", -1, month_after)
					Else
						next_paydate = next_pay_month & "/" & bimonthly_second & "/" & next_pay_year
					End If
				Else
					next_pay = DateAdd("m", 1, next_paydate)                                                            'go to the next month
					next_pay_month = DatePart("m", next_pay)
					next_pay_year = DatePart("yyyy", next_pay)
					next_paydate = next_pay_month & "/" & bimonthly_first & "/" & next_pay_year   'then go to the second pay date
				End If
			ElseIf pay_freq = "3 - Every Other Week" Then
				next_paydate = DateAdd("d", 14, next_paydate)
			ElseIf pay_freq = "4 - Every Week" Then
				next_paydate = DateAdd("d", 7, next_paydate)
			End If
			counter = counter + 1
		Loop until DateDiff("d", last_check, next_paydate) > 0           'We go until the loop has moved past the last pay date entered

		If left(list_of_all_paydates_start_to_finish, 1) = "~" Then     'now we make the list an array
			list_of_all_paydates_start_to_finish = right(list_of_all_paydates_start_to_finish, len(list_of_all_paydates_start_to_finish) - 1)
		End If

		If Instr(list_of_all_paydates_start_to_finish, "~") = 0 Then
			temp_expected_check_array = array(list_of_all_paydates_start_to_finish)
		Else
			temp_expected_check_array = split(list_of_all_paydates_start_to_finish, "~")
		End If

		If temp_expected_check_array(UBound(temp_expected_check_array)) <> last_check Then     'this got a little weird sometimes so it is just a double check
			temp_expected_check_array = ""
			list_of_all_paydates_start_to_finish = list_of_all_paydates_start_to_finish & "~" & next_paydate
			temp_expected_check_array = split(list_of_all_paydates_start_to_finish, "~")
		End If

		ReDim expected_check_array(UBound(temp_expected_check_array))
		For the_thing = 0 to UBound(temp_expected_check_array)
			expected_check_array(the_thing) = temp_expected_check_array(the_thing)
		Next
	end sub

    public sub create_months_check_list(update_month)
        'List of months updated

        checks_list = ""

        'Here we are making a list of all the checks that we expect for the active month - we will then make that an array
        If pay_freq = "1 - One Time Per Month" Then
            next_date = DateAdd("d", 1, first_check)
            If DatePart("d", next_date) = 1 Then
                first_of_mx_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
                first_of_next_month = DateAdd("m", 1, first_of_mx_month)
                the_day_of_pay = DateAdd("d", -1, first_of_next_month)
            Else
                day_of_month = DatePart("d", first_check)

                the_day_of_pay = MAXIS_footer_month & "/" & day_of_month & "/" & MAXIS_footer_year
                the_day_of_pay = DateValue(the_day_of_pay)
            End If
            If income_end_dt <> "" Then
                If DateDiff("d", the_day_of_pay, income_end_dt) >= 0 Then checks_list = checks_list & "~" & the_day_of_pay
            Else
                checks_list = checks_list & "~" & the_day_of_pay
            End If


        ElseIf pay_freq = "2 - Two Times Per Month" Then
            checks_in_month = 0
            If order_ubound <> "" Then
                For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
                    For all_income = 0 to UBound(pay_date)                  'then loop through all of the income information
                        'conditional if it is the right panel AND the order matches - then do the thing you need to do
                        If check_order(all_income) = order_number Then
                            If DatePart("m", pay_date(all_income)) = DatePart("m", update_month) AND DatePart("yyyy", pay_date(all_income)) = DatePart("yyyy", update_month) Then
                                checks_in_month = checks_in_month + 1
                                checks_list = checks_list & "~" & view_pay_date(all_income)
                            End If
                        End If
                    Next
                Next
            End If

            If checks_in_month = 0 Then
                month_to_use = DatePart("m", update_month)
                year_to_use = DatePart("yyyy", update_month)

                checks_list = checks_list & "~" & DateValue(month_to_use & "/" & bimonthly_first & "/" & year_to_use)
                If bimonthly_second = "LAST" Then
                    first_of_payMonth = month_to_use & "/1/" & year_to_use
                    first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                    checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                Else
                    checks_list = checks_list & "~" & DateValue(month_to_use & "/" & bimonthly_second & "/" & year_to_use)
                End If

            ElseIf checks_in_month = 1 Then
                the_check = replace(checks_list, "~", "")
                month_to_use = DatePart("m", update_month)
                year_to_use = DatePart("yyyy", update_month)
                If bimonthly_second = "LAST" Then
                    first_of_payMonth = month_to_use & "/1/" & year_to_use
                    first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                    If DatePart("d", the_check) = bimonthly_first Then
                        checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                    Else
                        the_other_check = month_to_use & "/" & bimonthly_first & "/" & year_to_use
                        checks_list = the_other_check & "~" & the_check
                    End If
                Else
                    If DatePart("d", the_check) = bimonthly_first Then
                        the_other_check = month_to_use & "/" & bimonthly_second & "/" & year_to_use
                        checks_list = checks_list & "~" & the_other_check
                    ElseIf DatePart("d", the_check) = bimonthly_second Then
                        the_other_check = month_to_use & "/" & bimonthly_first & "/" & year_to_use
                        checks_list = the_other_check & "~" & the_check
                    End If
                End If
            End If
        ElseIf pay_freq = "3 - Every Other Week" Then
            the_date = DateValue(first_check)
            Do
                If DatePart("m", the_date) = DatePart("m", update_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", update_month) Then
                    If income_end_dt <> "" Then
                        If DateDiff("d", the_date, income_end_dt) >= 0 Then checks_list = checks_list & "~" & the_date
                    Else
                        checks_list = checks_list & "~" & the_date
                    End If
                End If
                the_date = DateAdd("d", 14, the_date)
            Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr
        ElseIf pay_freq = "4 - Every Week" Then
            the_date = DateValue(first_check)
            Do
                If DatePart("m", the_date) = DatePart("m", update_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", update_month) Then
                    If income_end_dt <> "" Then
                        If DateDiff("d", the_date, income_end_dt) >= 0 Then checks_list = checks_list & "~" & the_date
                    Else
                        checks_list = checks_list & "~" & the_date
                    End If
                End If
                the_date = DateAdd("d", 7, the_date)
            Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr

        ElseIf pay_freq = "5 - Other" Then
        End If

        'formatting the list and maing it an array
        If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
        If InStr(checks_list, "~") <> 0 Then
            If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
            temp_array = Split(checks_list,"~")
        Else
            temp_array = Array(checks_list)
        End If
        ReDim this_month_checks_array(UBound(temp_array))
        for cow = 0 to UBound(temp_array)
            this_month_checks_array(cow) = temp_array(cow)
        next

        If pay_freq = "3 - Every Other Week" or pay_freq = "4 - Every Week" Then
            the_date = DateValue(first_check)
            beg_of_retro_month = RETRO_footer_month & "/1/" & RETRO_footer_year
            If DateDiff("d", first_check, beg_of_retro_month) < 0 Then
                If pay_freq = "3 - Every Other Week" Then list_start_date = DateAdd("d", -14, first_check)
                If pay_freq = "4 - Every Week" Then list_start_date = DateAdd("d", -7, first_check)
                Do While DateDiff("d", list_start_date, beg_of_retro_month) < 0
                    If pay_freq = "3 - Every Other Week" Then list_start_date = DateAdd("d", -14, list_start_date)
                    If pay_freq = "4 - Every Week" Then list_start_date = DateAdd("d", -7, list_start_date)
                Loop
                the_date = list_start_date
            End If
        End If

        checks_list = ""
        If pay_freq = "1 - One Time Per Month" Then
            next_date = DateAdd("d", 1, first_check)
            If DatePart("d", next_date) = 1 Then
                first_of_mx_month = RETRO_footer_month & "/1/" & RETRO_footer_year
                first_of_next_month = DateAdd("m", 1, first_of_mx_month)
                the_day_of_pay = DateAdd("d", -1, first_of_next_month)
            Else
                day_of_month = DatePart("d", first_check)

                the_day_of_pay = RETRO_footer_month & "/" & day_of_month & "/" & RETRO_footer_year
                the_day_of_pay = DateValue(the_day_of_pay)
            End If
            checks_list = checks_list & "~" & the_day_of_pay

        ElseIf pay_freq = "2 - Two Times Per Month" Then
            checks_in_month = 0
            If order_ubound <> "" Then
                For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
                    For all_income = 0 to UBound(pay_date)                  'then loop through all of the income information
                        'conditional if it is the right panel AND the order matches - then do the thing you need to do
                        If check_order(all_income) = order_number Then
                            If DatePart("m", pay_date(all_income)) = DatePart("m", RETRO_month) AND DatePart("yyyy", pay_date(all_income)) = DatePart("yyyy", RETRO_month) Then
                                checks_in_month = checks_in_month + 1
                                checks_list = checks_list & "~" & view_pay_date(all_income)
                            End If
                        End If
                    Next
                Next
            End If

            If checks_in_month = 0 Then
                checks_list = checks_list & "~" & DateValue(RETRO_footer_month & "/" & bimonthly_first & "/" & RETRO_footer_year)
                If bimonthly_second = "LAST" Then
                    first_of_payMonth = RETRO_footer_month & "/1/" & RETRO_footer_year
                    first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                    checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                Else
                    checks_list = checks_list & "~" & DateValue(RETRO_footer_month & "/" & bimonthly_second & "/" & RETRO_footer_year)
                End If

            ElseIf checks_in_month = 1 Then
                the_check = replace(checks_list, "~", "")
                If bimonthly_second = "LAST" Then
                    first_of_payMonth = RETRO_footer_month & "/1/" & RETRO_footer_year
                    first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
                    If DatePart("d", the_check) = bimonthly_first Then
                        checks_list = checks_list & "~" & DateAdd("d", -1, first_of_nextMonth)
                    Else
                        the_other_check = RETRO_footer_month & "/" & bimonthly_first & "/" & RETRO_footer_year
                        checks_list = the_other_check & "~" & the_check
                    End If
                Else
                    If DatePart("d", the_check) = bimonthly_first Then
                        the_other_check = RETRO_footer_month & "/" & bimonthly_second & "/" & RETRO_footer_year
                        checks_list = checks_list & "~" & the_other_check
                    ElseIf DatePart("d", the_check) = bimonthly_second Then
                        the_other_check = RETRO_footer_month & "/" & bimonthly_first & "/" & RETRO_footer_year
                        checks_list = the_other_check & "~" & the_check
                    End If
                End If
            End If

        ElseIf pay_freq = "3 - Every Other Week" Then
            Do
                If DatePart("m", the_date) = DatePart("m", RETRO_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", RETRO_month) Then
                    checks_list = checks_list & "~" & the_date
                End If
                the_date = DateAdd("d", 14, the_date)
            Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr
        ElseIf pay_freq = "4 - Every Week" Then
            Do
                If DatePart("m", the_date) = DatePart("m", RETRO_month) AND DatePart("yyyy", the_date) = DatePart("yyyy", RETRO_month) Then
                    checks_list = checks_list & "~" & the_date
                End If
                the_date = DateAdd("d", 7, the_date)
            Loop until right("0" & DatePart("m", the_date), 2) = CM_plus_2_mo AND right(DatePart("yyyy", the_date), 2) = CM_plus_2_yr
        ElseIf pay_freq = "5 - Other" Then
        End If

        'formatting and making this an array
        If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
        If InStr(checks_list, "~") <> 0 Then
            If left(checks_list, 1) = "~" Then checks_list = right(checks_list, len(checks_list)-1)
            temp_array = Split(checks_list,"~")
        Else
            temp_array = Array(checks_list)
        End If
        ReDim retro_month_checks_array(UBound(temp_array))
        for cow = 0 to UBound(temp_array)
            retro_month_checks_array(cow) = temp_array(cow)
        next

    end sub

	public sub create_new_panel()
        'If a new panel is to be created - this displays the dialog to collect the information and then creates the panel

		'NEW JOB PANEL Dialog'
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 431, 115, "New JOBS Panel"
			EditBox 75, 10, 20, 15, enter_JOBS_clt_ref_nbr
			DropListBox 155, 10, 60, 45, "W - Wages (Incl Tips)"+chr(9)+"J - WIOA"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program"+chr(9)+"N - Census Income", enter_JOBS_inc_type_code
			DropListBox 330, 10, 95, 45, ""+chr(9)+"01 - Subsidized Public Sector Employer"+chr(9)+"02 - Subsidized Private Sector Employer"+chr(9)+"03 - On-The-Job Training"+chr(9)+"04 - AmeriCorps(VISTA/State/National/NCCC)", enter_JOBS_subsdzd_inc_type
			DropListBox 155, 30, 90, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd"+chr(9)+"? - Unknown", enter_JOBS_verif_code
			EditBox 330, 30, 50, 15, enter_JOBS_hrly_wage
			EditBox 155, 50, 195, 15, enter_JOBS_employer
			EditBox 155, 70, 50, 15, enter_JOBS_start_date
			EditBox 330, 70, 50, 15, enter_JOBS_end_date
			CheckBox 105, 95, 30, 10, "SNAP", apply_to_SNAP
			CheckBox 145, 95, 30, 10, "CASH", apply_to_CASH
			CheckBox 190, 95, 20, 10, "HC", apply_to_HC
			CheckBox 230, 95, 30, 10, "GRH", apply_to_GRH
			ButtonGroup ButtonPressed
				OkButton 320, 95, 50, 15
				CancelButton 375, 95, 50, 15
			Text 10, 15, 65, 10, "Client Ref Number:"
			Text 105, 15, 45, 10, "Income Type:"
			Text 240, 15, 85, 10, "Subsidized Income Type:"
			Text 110, 35, 40, 10, "Verification:"
			Text 280, 35, 50, 10, "Hourly Wage:"
			Text 115, 55, 35, 10, "Employer:"
			Text 105, 75, 45, 10, "Income Start:"
			Text 285, 75, 40, 10, "Income End:"
			Text 10, 95, 90, 10, "Apply Income to Programs:"
		EndDialog

		cancel_clarify = ""         'blanking this out from previous dialog or another loop
		panel_created = TRUE     'defaulting to having the panel created
		Do
			Do
				err_msg = ""

				dialog Dialog1

				'alternate for cancel_confirmation
				If ButtonPressed = 0 then       'this is the cancel button
					cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no actions taken so far will be noted.", vbQuestion + vbYesNo, "Clarify Cancel")
					If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")     'ends the script entirely
					'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
                    If cancel_clarify = vbNo Then       'cancels the current operation without cancelling the script
                        panel_created = FALSE        'this keeps a blank panel from being created if 'Cancel' is selected
                        Exit Do
                    End If
				End if

				'Error handling
				enter_JOBS_employer = trim(enter_JOBS_employer)
				If trim(enter_JOBS_clt_ref_nbr) = "" Then err_msg = err_msg & vbNewLine & "* Enter the member number of the client that is employed at this job."               'Need a member reference number
				If len(enter_JOBS_clt_ref_nbr) <> 2 Then err_msg = err_msg & vbNewLine & "* The member number should be 2 digits"                                               'Ensuring the member reference number is accurate
				If trim(enter_JOBS_inc_type_code) = "  " Then err_msg = err_msg & vbNewLine & "* Enter the income type of the job."                                             'Need JOB income type - usually 'W'
				If trim(enter_JOBS_verif_code) = "  " Then err_msg = err_msg & vbNewLine & "* Enter the verification code for this job."                                        'MAXIS requires some kind of verification code - 'N' is okay here
				If enter_JOBS_employer = "" Then err_msg = err_msg & vbNewLine & "* Enter the employer name for this job."                                                      'Must have an employer name to enter into MAXIS
				If len(enter_JOBS_employer) > 30 Then err_msg = err_msg & vbNewLine & "* The Employer name is too long to fit on the JOBS panel, abbreviate as necessary."      'Employer name line on JOBS is only a certain length, conforming to that length
				If IsDate(enter_JOBS_start_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for Income start date."                                     'Requiring a start date for income since SNAP eligibility results can only be created if we have this
				If trim(enter_JOBS_end_date) <> "" AND IsDate(enter_JOBS_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid date for Income end date."     'IF there is an end date, it must be a valid date

				if err_msg <> "" Then msgBox "Please resolve the following to continue:" & vbNewLine & err_msg          'displaying the error messages if there are any

			Loop until err_msg = ""
			call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
		LOOP UNTIL are_we_passworded_out = false
		member = enter_JOBS_clt_ref_nbr

		If panel_created = TRUE Then                     'only continues if cancel was not selected above
			info_saved = False
			Call navigate_to_MAXIS_screen("CASE", "CURR")   'We need the most recent case application date because we cannot go further back than that month/year

			EMReadScreen appl_date, 8, 8, 29                'this is where the case application date is - NOT program specific

			beginning_month = ""
			beginning_year = ""

			If enter_JOBS_clt_ref_nbr <> "01" Then
				MAXIS_footer_month = CM_mo
				MAXIS_footer_year = CM_yr

				Call navigate_to_MAXIS_screen("STAT", "MEMB")
				Call write_value_and_transmit(enter_JOBS_clt_ref_nbr, 20, 76)

				EMReadScreen memb_arrival_date, 8, 4, 73
				If memb_arrival_date <> "        " Then
					memb_arrival_date = replace(memb_arrival_date, " ", "/")
					memb_arrival_date = DateAdd("d", 0, memb_arrival_date)

					beginning_month = DatePart("m", memb_arrival_date)                  'we use the application date as the footer month and year to enter the JOBS panel
					beginning_year = DatePart("yyyy", memb_arrival_date)
					first_check = beginning_month & "/01/" & beginning_year         'setting the date of the first check to be entered on JOBS
				End If
			End If

			If beginning_month = "" Then
				If DateDiff("m", appl_date, enter_JOBS_start_date) >= 0 Then    'if the application date is before the income start date
					beginning_month = DatePart("m", enter_JOBS_start_date)      'we use the income start date as the footer month and year to enter the JOBS panel
					beginning_year = DatePart("yyyy", enter_JOBS_start_date)
					first_check = enter_JOBS_start_date                             'setting the date of the first check to be entered on JOBS
				Else                                                            'otherwise, if the job started before the application date
					beginning_month = DatePart("m", appl_date)                  'we use the application date as the footer month and year to enter the JOBS panel
					beginning_year = DatePart("yyyy", appl_date)
					first_check = beginning_month & "/01/" & beginning_year         'setting the date of the first check to be entered on JOBS
				End If
			End If

			beginning_month = right("00"&beginning_month, 2)                'creating 2 digit month and year variables
			beginning_year = right(beginning_year, 2)

			If DateDiff("m", first_check, date) > 12 Then                   'if the first check to be entered on the panel is more that 12 months from the current date
																			'script will confirm the month and year to add the panel in MAXIS
				'PROCEDURE CLARIFICATION - allowing workers to adjust the month and year the panel is entered
				'CONFIRM ADD PANEL MONTH Dialog'
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 191, 175, "Confirm Update Month"
					EditBox 140, 60, 15, 15, beginning_month
					EditBox 160, 60, 15, 15, beginning_year
					ButtonGroup ButtonPressed
						OkButton 135, 155, 50, 15
					Text 10, 10, 165, 10, "** This new job started at least 12 months ago. **"
					Text 10, 30, 165, 20, "The script will go back to " & beginning_month & "/" & beginning_year & " to add this JOBS panel in the month the job started."
					Text 10, 60, 120, 15, "If this needs to be adjusted, change the footer month and year here:"
					GroupBox 10, 85, 170, 65, "Info"
					Text 20, 100, 150, 40, "Best practice is to add the job information in the footer month and year the income started. An exception may be if the job was currently in STAT and the panel was deleted, only add the information in the first month of deletion."
				EndDialog

				dialog Dialog1

				beginning_month = beginning_month * 1       'making these numbers for error handling
				beginning_year = beginning_year * 1

				If beginning_month = DatePart("m", enter_JOBS_start_date) AND beginning_year = DatePart("yyyy", enter_JOBS_start_date) Then     'setting the first check to be entered on JOBS
					first_check = enter_JOBS_start_date
				Else
					first_check = beginning_month & "/01/" & beginning_year
				End If

				beginning_month = right("00"&beginning_month, 2)            'making these 2-digit variables
				beginning_year = right(beginning_year, 2)

			End If

			MAXIS_footer_month = beginning_month        'setting the footer month and year to the update month so nav functions work
			MAXIS_footer_year = beginning_year

			Call back_to_SELF           'Getting out of STAT to be sure footer month is correct

			Call navigate_to_MAXIS_screen("STAT", "SUMM")
            pay_placeholder = first_check

			Do                          'this loop is to update future months with the JOB information
				STATS_manualtime = STATS_manualtime + 85
				If info_saved = FALSE Then				'If the information has not yet been saved to the array it means we are in the first month
					EMWriteScreen "JOBS", 20, 71		'go to JOBS
					EMWriteScreen member, 20, 76		'go to the right member
					EMWriteScreen "NN", 20, 79			'create new JOBS panel

					transmit
				Else                                	'If the information is in the array, we will use that to navigate
					EMWriteScreen "JOBS", 20, 71        'go to JOBS
					EMWriteScreen member, 20, 76		'here we use 'the_panel-1' because it would have been incremented on the previous loop after saving the information
					EMWriteScreen instance, 20, 79

					transmit

					EMReadScreen check_for_panel, 14, 24, 13                'sometimes the panel does not exist in a future month because data expires, we then need to add it again
					If check_for_panel = "DOES NOT EXIST" Then
						EMWriteScreen "JOBS", 20, 71
						EMWriteScreen member, 20, 76
						EMWriteScreen "NN", 20, 79

						transmit
					Else
						PF9             'other wise put it in EDIT MODE
					End If
				End If

				EMWriteScreen left(enter_JOBS_inc_type_code, 1), 5, 34      'adding all the information from the dialog - using only the codes which are stored in the leftmost portion of the variables
				EMWriteScreen left(enter_JOBS_subsdzd_inc_type, 2), 5, 74
				EMWriteScreen left(enter_JOBS_verif_code, 1), 6, 34
				EMWriteScreen "      ", 6, 75
				EMWriteScreen enter_JOBS_hrly_wage, 6, 75
				EMWriteScreen enter_JOBS_employer, 7, 42

				Call write_date(enter_JOBS_start_date, "MM DD YY", 9, 35)   'entering the start date
				If trim(enter_JOBS_end_date) <> "" Then                     'TESTING NEEDED - I do not believe this functionality has been well used in testing
					Call write_date(enter_JOBS_end_date, "MM DD YY", 9, 49) 'entering the end date if one was listed
					If DateDiff("d", pay_placeholder, enter_JOBS_end_date) >= 0 Then    'as long as the end date is after the date of the check to entering - the ceck is entered with $0 pay amount
						Call write_date(pay_placeholder, "MM DD YY", 12, 54)
						EMWriteScreen "    0.00", 12, 67
						EMWriteScreen "0  ", 18, 72
					Else                                                    'otherwise the pay information is blanked out
						EMWriteScreen "  ", 12, 54
						EMWriteScreen "  ", 12, 57
						EMWriteScreen "  ", 12, 60
						EMWriteScreen "        ", 12, 67
						EMWriteScreen "   ", 18, 72
					End If
				Else                                                        'if there is no end date entered, the first pay date and $0 pay is entered
					Call write_date(pay_placeholder, "MM DD YY", 12, 54)
					EMWriteScreen "    0.00", 12, 67
					EMWriteScreen "0  ", 18, 72
				End If
				transmit
				EMReadScreen check_for_error_prone_warning, 20, 6, 43       'some times some warnings come up - need to move from them
				If check_for_error_prone_warning = "Error Prone Warnings" Then transmit

                EMReadScreen new_panel, 1, 2, 73            'reading the panel instance that was created. It is only generated AFTER transmit has saved the entered information

				If info_saved = FALSE Then                                  'If we have not already saved the information to EARNED_INCOME_PANELS_ARRAY - this will do it here

					instance = "0" & new_panel
					income_received = False
					initial_month_mo = MAXIS_footer_month    'defaulting the first date to update to the month/year the panel was created
					initial_month_yr = MAXIS_footer_year

					Call read_panel				'capture current panel information

					info_saved = TRUE               'changing this variable for the next loop
				End If
                new_panel = True
				'Navigates to the current month + 1 footer month without sending the case through background
				CALL write_value_and_transmit("BGTX", 20, 71)
				CALL write_value_and_transmit("y", 16, 54)

				EMReadScreen all_months_check, 24, 24, 2    'this reads the error message at the bottom of STAT/WRAP if we cannot get to the next month because we are in CM+1
				EMReadScreen MAXIS_footer_month, 2, 20, 55  'If we are successful in getting to the next month, the footer month and year are set here
				EMReadScreen MAXIS_footer_year, 2, 20, 58

				pay_placeholder = MAXIS_footer_month & "/01/" & MAXIS_footer_year   'need a check date in the current footer month to enter on JOBS
			Loop until all_months_check = "CONTINUATION NOT ALLOWED"
			PF3 'leaving STAT - sending the case through background
		End If      'If panel_created = TRUE Then'
	end sub

    public sub confirm_job_budget()
        'displays the budget details and cacluations for the worker to confirm

        update_bimonthly_pay_dates_btn			= 200
		back_to_checks_btn						= 210
		budget_correct_btn						= 220

		pay_frequency_tips_and_tricks_btn 		= 1001
        cm_budg_btn                             = 1002
		confirm_snap_budget_tips_and_tricks_btn = 1003
		confirm_cash_budget_tips_and_tricks_btn = 1004
		hc_retro_budget_tips_and_tricks_btn 	= 1005
		confirm_hc_budget_tips_and_tricks_btn 	= 1006

		Do
			err_msg = ""

			dlg_len = 120        'starting with this dialog

			If pay_freq = "2 - Two Times Per Month" Then
				dlg_len = dlg_len + 20
			End If
			If apply_to_SNAP = checked Then       'resizing the dialog and the SNAP Groupbox if income applies to SNAP
                grp_len = 40
                'adjust sizing based on number of checks
                grp_len = grp_len + (snap_check_count * 10) + 15
                If grp_len < 80 Then grp_len = 80

				dlg_len = dlg_len + grp_len + 10
			End If
			If verif_type = "? - EXPEDITED SNAP ONLY" Then      'adding size for XFS information tot be added to dialog
				dlg_len = dlg_len + 40
			End If
			If apply_to_CASH = checked or apply_to_GRH = checked Then       'resizing the dialog and the Cash Groupbox if income applies to Cash
                cash_grp_len = 60
                'adjust sizing based on number of checks
                cash_grp_len = cash_grp_len + (cash_check_count * 10) + 15
                If cash_grp_len < 100 Then cash_grp_len = 100

                dlg_len = dlg_len + cash_grp_len + 10
                If apply_to_SNAP = checked Then dlg_len = dlg_len - 5
			End If
			If apply_to_HC = checked Then         'resizing the dialog and the HC Groupbox if income applies to HC
				dlg_len = dlg_len + 5
				hc_grp_len = 60
				length_of_checks_list = total_check_count*10
				If length_of_checks_list < 60 Then length_of_checks_list = 60

				hc_grp_len = hc_grp_len + length_of_checks_list
				dlg_len = dlg_len + hc_grp_len
				If length_of_checks_list < 60 Then dlg_len = dlg_len + 5
			End If

			y_pos = 25      'incrementer to move things down

			'CONFIRM BUDGET Dialog - mostly shows the information after being calculated for each program and makes the worker confirm this is correct
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 421, dlg_len, "Confirm JOBS Budget for " & employer
				Text 10, 10, 100, 20, "JOBS " & member & " " & instance & " - " & employer '& "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
				Text 115, 10, 140, 10, "Pay Frequency - " & pay_freq '& "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
				CheckBox 260, 10, 130, 10, "Check here to confirm pay frequency.", confirm_pay_freq_checkbox
				If pay_freq = "2 - Two Times Per Month" Then
					If bimonthly_second <> "LAST" Then Text 15, y_pos+5, 150, 10, "BiMonthly Pay Dates are: " & bimonthly_first & " and " & bimonthly_second & "."
					If bimonthly_second = "LAST" Then Text 15, y_pos+5, 210, 10, "BiMonthly Pay Dates are: " & bimonthly_first & " and the LAST day of the month."
					ButtonGroup ButtonPressed
						PushButton 210, y_pos, 115, 15, "Change Semi-Monthly Dates", update_bimonthly_pay_dates_btn
					y_pos = y_pos + 20
				End If

				If verif_type = "? - EXPEDITED SNAP ONLY" Then
					Text 10, y_pos, 400, 10, "THIS INCOME HAS NOT BEEN VERIFIED - '?' verification code used."
					Text 10, y_pos +10, 400, 10, " -- Only SNAP can be handled this way. The script will only apply SNAP budgeting functionality.-- "
					Text 10, y_pos + 20, 400, 10, "A note will be added that some or all of pay information is only reported by client and not verified."
					y_pos = y_pos + 40
				End If

				Text 20, y_pos + 5, 290, 10, "*** ALL PROGRAM BUDGETS NEED TO BE CONFIRMED ***       "
				Text 15, y_pos + 15 , 290, 10, "Review details and confirm by using the checkboxes to continue."
                Text 260, y_pos, 170, 10, "Hourly Wage: $ " & hourly_wage
                Text 260, y_pos + 10, 170, 10, "Average Hours per Check: " & ave_hrs_per_pay
                Text 260, y_pos + 20, 170, 10, "Average Paycheck: $ " & ave_inc_per_pay '& "   -   Monthly Income: $ " & monthly_income
				y_pos = y_pos + 35
				' y_pos = y_pos + 15

                Text 10, y_pos, 215, 10, " * ! * ! * CHECK CM 22.03.01    FOR BUDGETING POLICY * ! * ! * "
                cm_budg_y_pos = y_pos

                If pick_one = use_actual Then Text 225, y_pos, 200, 10, "Income provided covers the period " & first_check & " to " & last_check & "."
				If pick_one = use_estimate Then Text 225, y_pos, 200, 10, "Income is based on anticipated hours and rate of pay."
				y_pos = y_pos + 15

				If apply_to_SNAP = checked Then
					GroupBox 5, y_pos, 410, grp_len, "SNAP Budget - ENTERED ON SNAP PIC"
					y_pos = y_pos + 10
					check_list_y_pos = y_pos

					list_pos = 0      'multiplier to move the array items down
					If pick_one = use_actual Then
						' 'this part actually looks at the income information IN ORDER
						For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
							For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
								'conditional if it is the right panel AND the order matches - then do the thing you need to do
								If check_order(all_income) = order_number and SNAP_dialog_display(all_income) <> "" Then
									Text 20, (list_pos * 10) + y_pos + 10, 240, 10, view_pay_date(all_income) & " - $ " & SNAP_dialog_display(all_income)
									list_pos = list_pos + 1
								End If
							next
						next
						If SNAP_list_of_excluded_pay_dates = "" Then
							Text 20, (list_pos * 10) + y_pos + 15, 140, 10, "~ No Excluded Checks ~"
                        Else
                            Text 10, (list_pos * 10) + y_pos + 15, 240, 20, "Paychecks not included: " & SNAP_list_of_excluded_pay_dates      'list of all excluded pay dates
						End If
						list_pos = list_pos + 2
					ElseIf pick_one = use_estimate Then
						For each money_day in snap_anticipated_pay_array      'this is the list we made above - it has pay date and amount in each item
							Text 20, ((list_pos+1) * 10) + y_pos, 90, 10, money_day
							list_pos = list_pos + 1
						Next
					End If
					GroupBox 10, check_list_y_pos, 250, (list_pos*10) + 10, paycheck_list_title        '"Paychecks Inclued in Budget:"'

					Text 270, y_pos, 140, 10, 		"Average hourly rate of pay:   $ " & snap_hourly_wage
					Text 277, y_pos + 13, 135, 10, 	"Average " & word_for_freq & " hours:     " & snap_ave_hrs_per_pay
					Text 270, y_pos + 26, 140, 10, 	"Average paycheck amount:   $ " & snap_ave_inc_per_pay
					Text 272, y_pos + 39, 140, 10, 	"Monthly Budgeted Income:   $ " & SNAP_monthly_income

                    orig_y_pos = y_pos
                    y_pos = y_pos + ((list_pos) * 10)+10
                    If y_pos - orig_y_pos < 55 Then y_pos = orig_y_pos + 55

                    CheckBox 10, y_pos, 330, 10, "Check here to confirm this SNAP budget is correct and is the best estimate of anticipated income.", SNAP_accurate_checkbox
					csbtnt_y_pos = y_pos-5
					y_pos = y_pos + 20
				End If        'If apply_to_SNAP = checked Then

				If apply_to_CASH = checked or apply_to_GRH = checked Then
					GroupBox 5, y_pos, 410, cash_grp_len, "CASH / GRH Budget - ENTERED ON CASH PIC"
                    y_pos = y_pos + 10
					Text 10, y_pos + 5, 220, 10, "If this income is excluded from the Cash budget, select the reason:"
					DropListBox 230, y_pos, 160, 15, "NONE"+chr(9)+"Caregiver under 20 - 50% in school"+chr(9)+"Child under 18 in school"+chr(9)+"Excluded Work Program"+chr(9)+"Excluded Spousal Income", income_excluded_cash_reason
					y_pos = y_pos + 20

					check_list_y_pos = y_pos
					list_pos = 0      'multiplier to move the array items down
					If pick_one = use_actual Then
						' 'this part actually looks at the income information IN ORDER
						For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
							For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
								'conditional if it is the right panel AND the order matches - then do the thing you need to do
								If check_order(all_income) = order_number and CASH_dialog_display(all_income) <> "" Then
									Text 20, (list_pos * 10) + y_pos + 10, 240, 10, view_pay_date(all_income) & " - $ " & CASH_dialog_display(all_income)
									list_pos = list_pos + 1
								End If
							next
						next
						If CASH_list_of_excluded_pay_dates = "" Then
							Text 20, (list_pos * 10) + y_pos + 15, 140, 10, "~ No Excluded Checks ~"
                        Else
                            Text 10, (list_pos * 10) + y_pos + 15, 240, 20, "Paychecks not included: " & CASH_list_of_excluded_pay_dates      'list of all excluded pay dates
						End If
						list_pos = list_pos + 2
					ElseIf pick_one = use_estimate Then
						For each money_day in cash_anticipated_pay_array      'this is the list we made above - it has pay date and amount in each item
							Text 20, (list_pos * 10) + y_pos + 10, 90, 10, money_day
							list_pos = list_pos + 1
						Next
					End If
					GroupBox 10, check_list_y_pos, 250, (list_pos*10) + 10, paycheck_list_title        '"Paychecks Inclued in Budget:"'

					Text 270, y_pos, 140, 10, 		"Average hourly rate of pay:   $ " & cash_hourly_wage
					Text 277, y_pos + 13, 135, 10, 	"Average " & word_for_freq & " hours:      " & cash_ave_hrs_per_pay
					Text 270, y_pos + 26, 140, 10, 	"Average paycheck amount:   $ " & cash_ave_inc_per_pay
					Text 272, y_pos + 39, 140, 10, 	"Monthly Budgeted Income:   $ " & CASH_monthly_income

                    orig_y_pos = y_pos
                    y_pos = y_pos + ((list_pos) * 10)+10
                    If y_pos - orig_y_pos < 55 Then y_pos = orig_y_pos + 55

					CheckBox 10, y_pos, 330, 10, "Check here to confirm this CASH budget is correct and is the best estimate of anticipated income.", CASH_accurate_checkbox
					ccbtnt_y_pos = y_pos - 5
					y_pos = y_pos + 20
				End If            'If apply_to_CASH = checked Then

				If apply_to_HC = checked Then
					GroupBox 5, y_pos, 410, hc_grp_len, "Health Care Budget"
					y_pos = y_pos + 15
					Text 10, y_pos, 400, 10, "Pay information will be entered on the prospective side only, using actual or estimated pay amounts."
					y_pos = y_pos + 10
					Text 10, y_pos, 30, 10, "CHECKS"

					y_pos = y_pos + 10
					Text 150, y_pos, 250, 10, "Average amount per pay period: $" & ave_inc_per_pay & " - for HC Inc Est Pop-up."

					Text 150, y_pos + 15, 200,10, "Notes about HC Budget:"
					EditBox 150, y_pos + 25, 250, 15, hc_budg_notes

					CheckBox 150, y_pos + 45, 170, 10, "Check here if HC needs a Retrospective Budget.", HC_RETRO_accurate_checkbox
					hcrtnt_y_pos = y_pos + 40
					list_pos = 0
					' 'this part actually looks at the income information IN ORDER
					For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
						For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
							'conditional if it is the right panel AND the order matches - then do the thing you need to do
							If check_order(all_income) = order_number Then
								Text 20, (list_pos * 10) + y_pos, 125, 10, view_pay_date(all_income) & " - $" & gross_amount(all_income) & " - " & hours(all_income) & "hrs."
								list_pos = list_pos + 1
							End If
						next
					next

					bottom_of_checks = y_pos + (list_pos * 10)
					If list_pos < 6 Then bottom_of_checks = y_pos + 60
					y_pos = bottom_of_checks + 5

					CheckBox 10, y_pos, 230, 10, "Check here if these checks and estimated pay amount are accurate.", HC_accurate_checkbox
					chbtnt_y_pos = y_pos -5
					y_pos = y_pos + 20
				End If

				y_pos = dlg_len - 50
				GroupBox 5, y_pos, 410, 30, "Budget Explanation Conversation Details:"
				Text 10, y_pos + 15, 60, 10, "Conversation with:"
				ComboBox 75, y_pos + 10, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer"+chr(9)+spoke_with, spoke_with
				Text 140, y_pos + 15, 25, 10, "clarifies"
				EditBox 170, y_pos + 10, 235, 15, convo_detail
				y_pos = y_pos + 30

				ButtonGroup ButtonPressed
					PushButton 395, 5, 15, 15, "!", pay_frequency_tips_and_tricks_btn

					If apply_to_SNAP = checked Then PushButton 340, csbtnt_y_pos, 15, 15, "!", confirm_snap_budget_tips_and_tricks_btn
                    PushButton 60, cm_budg_y_pos-3, 47, 13, "CM 22.03.01", cm_budg_btn
					If apply_to_CASH = checked Then PushButton 340, ccbtnt_y_pos, 15, 15, "!", confirm_cash_budget_tips_and_tricks_btn
					If apply_to_HC = checked Then PushButton 320, hcrtnt_y_pos, 15, 15, "!", hc_retro_budget_tips_and_tricks_btn
					If apply_to_HC = checked Then PushButton 240, chbtnt_y_pos, 15, 15, "!", confirm_hc_budget_tips_and_tricks_btn

					PushButton 10, y_pos, 100, 15, "Go Back to Check List", back_to_checks_btn
					PushButton 265, y_pos, 100, 15, "Budget is Accurate", budget_correct_btn
					' OkButton 315, y_pos, 50, 15
					CancelButton 365, y_pos, 50, 15
			EndDialog

			Dialog Dialog1      'calling the dialog
			save_your_work

			If ButtonPressed = -1 Then ButtonPressed = budget_correct_btn
            If ButtonPressed = budget_correct_btn Then budget_confirmed = True

			If ButtonPressed = update_bimonthly_pay_dates_btn Then
				Call dialog_for_bimonthly_pay()
				err_msg = "LOOP"         'makes the dialog loop back without displaying an error message
			End If

			If apply_to_SNAP = unchecked Then SNAP_accurate_checkbox = checked
			If apply_to_CASH = unchecked and apply_to_GRH = unchecked Then CASH_accurate_checkbox = checked
			If apply_to_HC = unchecked Then HC_accurate_checkbox = checked

			If confirm_pay_freq_checkbox = unchecked Then err_msg = err_msg & vbCr & "* Review the pay frequency and confirm."
			If SNAP_accurate_checkbox = unchecked Then err_msg = err_msg & "*** SNAP Budget not confirmed. ***" & vbCr & " - Review details of SNAP budget, return to checks entry if inaccurate." & vbCr
			If CASH_accurate_checkbox = unchecked Then
                err_progs = ""
                If apply_to_CASH = checked Then err_progs = err_progs & "CASH "
                If apply_to_GRH = checked Then err_progs = err_progs & "GRH "
                err_progs = replace(trim(err_progs), " ", " and ")
                err_msg = err_msg & vbCr & "*** " & err_progs & " Budget not confirmed***" & vbCr & " - Review details of " & replace(err_progs, "and", "/") & " budget, return to checks entry if inaccurate." & vbCr
            End If
            If HC_accurate_checkbox = unchecked Then err_msg = err_msg & vbCr & "*** HC Budget not confirmed ***" & vbCr & " - Review details of HC budget, return to checks entry if inaccurate." & vbCr

			If ButtonPressed = pay_frequency_tips_and_tricks_btn		Then tips_and_tricks_msg = MsgBox(pay_freq_2_msg_text, vbInformation, "Tips and Tricks")
            If ButtonPressed = cm_budg_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_00220301"
			If ButtonPressed = confirm_snap_budget_tips_and_tricks_btn 	Then tips_and_tricks_msg = MsgBox(confirm_snap_msg_text, vbInformation, "Tips and Tricks")
			If ButtonPressed = confirm_cash_budget_tips_and_tricks_btn 	Then tips_and_tricks_msg = MsgBox(confirm_cash_msg_text, vbInformation, "Tips and Tricks")
			If ButtonPressed = hc_retro_budget_tips_and_tricks_btn 		Then tips_and_tricks_msg = MsgBox(hc_retro_msg_text, vbInformation, "Tips and Tricks")
			If ButtonPressed = confirm_hc_budget_tips_and_tricks_btn 	Then tips_and_tricks_msg = MsgBox(confirm_hc_msg_text, vbInformation, "Tips and Tricks")

			If ButtonPressed > 1000 Then err_msg = "LOOP"

			If ButtonPressed = back_to_checks_btn Then
				budget_confirmed = False
				err_msg = ""
			End If


			'if the 'Cancel' button is pressed, the worker gets 3 options 1. cancel script, 2, cancel current job update, 3. Ooops, pressed cancel by mistake
			If ButtonPressed = 0 then
				cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no information provided so far will be updated or noted. If you choose 'No' the update for THIS JOB will be cancelled and rest of the script will continue." & vbNewLine & vbNewLine & "YES - Stop the script entirely." & vbNewLine & "NO - Do not stop the script entrirely, just cancel the entry of this job information."& vbNewLine & "CANCEL - I didn't mean to cancel at all. (Cancel my cancel)", vbQuestion + vbYesNoCancel, "Clarify Cancel")
				If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")
				'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
				If cancel_clarify = vbNo Then           'this is to cancel the job update
					budget_confirmed = True
					income_received = False   'this makes the script skip this job in the next functions
					Exit Do
				End If          'there is no vbCancel handling because the script just continues at that point.
				If cancel_clarify = vbCancel Then err_msg = "LOOP"
			End if

			If err_msg <> "" and err_msg <> "LOOP" Then
				msg_display = "BUDGET FOR " & employer & " NOT CONFIRMED" & vbCr & vbCr
				msg_display = msg_display & "All of the information provided as checks and estimates need to be reviewed for each program in this dialog." & vbCr & vbCr
				msg_display = msg_display & "It appears not all details have been reviewed: " & err_msg
				MsgBox msg_display
			End If

		Loop until err_msg = ""
	end sub

	public sub delete_one_check(check_instance)
        'deletes a check from the list for this job

        If check_instance > UBound(pay_date) Then
			'do nothing - the check doesn't exist or would break everything
		ElseIf check_instance = 0 and UBound(pay_date) = 0 Then
			ReDim pay_date(0)
			ReDim gross_amount(0)
			ReDim hours(0)
			ReDim exclude_entirely(0)
			ReDim exclude_from_SNAP(0)
			ReDim exclude_from_CASH(0)
			ReDim reason_to_exclude(0)
			ReDim exclude_ALL_amount(0)
            ReDim exclude_ALL_hours(0)
			ReDim exclude_SNAP_amount(0)
            ReDim exclude_SNAP_hours(0)
			ReDim exclude_CASH_amount(0)
            ReDim exclude_CASH_hours(0)
			ReDim SNAP_info_string(0)
			ReDim CASH_info_string(0)
			ReDim check_order(0)
			ReDim view_pay_date(0)
			ReDim frequency_issue(0)
			ReDim future_check(0)
			ReDim duplicate_pay_date(0)
			ReDim reason_SNAP_amt_excluded(0)
			ReDim reason_CASH_amt_excluded(0)
			ReDim pay_detail_btn(0)
			ReDim check_info_entered(0)
			ReDim bonus_check(0)
			ReDim pay_split_regular_amount(0)
			ReDim pay_split_bonus_amount(0)
			ReDim pay_split_ot_amount(0)
            ReDim pay_split_ot_hours(0)
			ReDim pay_split_shift_diff_amount(0)
			ReDim pay_split_tips_amount(0)
			ReDim pay_split_other_amount(0)
			ReDim pay_split_other_detail(0)
			ReDim pay_excld_bonus(0)
			ReDim pay_excld_ot(0)
			ReDim pay_excld_shift_diff(0)
			ReDim pay_excld_tips(0)
			ReDim pay_excld_other(0)
			ReDim split_check_string(0)
			ReDim split_check_excld_string(0)
            ReDim split_exclude_amount(0)
			ReDim duplct_pay_date(0)
			ReDim calculated_by_ytd(0)
			ReDim ytd_calc_notes(0)
			ReDim pay_detail_exists(0)
			ReDim combined_into_one(0)
			ReDim SNAP_dialog_display(0)
			ReDim CASH_dialog_display(0)

		Else
			If check_instance <> UBound(pay_date) Then
				For check_item = 0 to UBound(pay_date)-1
					If check_item >= check_instance Then
						pay_date(check_item) 					= pay_date(check_item + 1)
						gross_amount(check_item) 				= gross_amount(check_item + 1)
						hours(check_item) 						= hours(check_item + 1)
						exclude_entirely(check_item) 			= exclude_entirely(check_item + 1)
						exclude_from_SNAP(check_item) 			= exclude_from_SNAP(check_item + 1)
						exclude_from_CASH(check_item) 			= exclude_from_CASH(check_item + 1)
						reason_to_exclude(check_item) 			= reason_to_exclude(check_item + 1)
						exclude_ALL_amount(check_item) 			= exclude_ALL_amount(check_item + 1)
                        exclude_ALL_hours(check_item)           = exclude_ALL_hours(check_item + 1)
						exclude_SNAP_amount(check_item) 		= exclude_SNAP_amount(check_item + 1)
                        exclude_SNAP_hours(check_item)          = exclude_SNAP_hours(check_item + 1)
						exclude_CASH_amount(check_item) 		= exclude_CASH_amount(check_item + 1)
                        exclude_CASH_hours(check_item)          = exclude_CASH_hours(check_item + 1)
						SNAP_info_string(check_item)			= SNAP_info_string(check_item + 1)
						CASH_info_string(check_item) 			= CASH_info_string(check_item + 1)
						check_order(check_item) 				= check_order(check_item + 1)
						view_pay_date(check_item) 				= view_pay_date(check_item + 1)
						frequency_issue(check_item) 			= frequency_issue(check_item + 1)
						future_check(check_item) 				= future_check(check_item + 1)
						duplicate_pay_date(check_item) 			= duplicate_pay_date(check_item + 1)
						reason_SNAP_amt_excluded(check_item) 	= reason_SNAP_amt_excluded(check_item + 1)
						reason_CASH_amt_excluded(check_item) 	= reason_CASH_amt_excluded(check_item + 1)
						pay_detail_btn(check_item) 				= pay_detail_btn(check_item + 1)
						check_info_entered(check_item) 			= check_info_entered(check_item + 1)
						bonus_check(check_item) 				= bonus_check(check_item + 1)
						pay_split_regular_amount(check_item) 	= pay_split_regular_amount(check_item + 1)
						pay_split_bonus_amount(check_item) 		= pay_split_bonus_amount(check_item + 1)
						pay_split_ot_amount(check_item) 		= pay_split_ot_amount(check_item + 1)
                        pay_split_ot_hours(check_item)          = pay_split_ot_hours(check_item + 1)
						pay_split_shift_diff_amount(check_item) = pay_split_shift_diff_amount(check_item + 1)
						pay_split_tips_amount(check_item) 		= pay_split_tips_amount(check_item + 1)
						pay_split_other_amount(check_item) 		= pay_split_other_amount(check_item + 1)
						pay_split_other_detail(check_item) 		= pay_split_other_detail(check_item + 1)
						pay_excld_bonus(check_item) 			= pay_excld_bonus(check_item + 1)
						pay_excld_ot(check_item) 				= pay_excld_ot(check_item + 1)
						pay_excld_shift_diff(check_item) 		= pay_excld_shift_diff(check_item + 1)
						pay_excld_tips(check_item) 				= pay_excld_tips(check_item + 1)
						pay_excld_other(check_item) 			= pay_excld_other(check_item + 1)
						split_check_string(check_item)			= split_check_string(check_item + 1)
                        split_check_excld_string(check_item)    = split_check_excld_string(check_item + 1)
                        split_exclude_amount(check_tem)         = split_exclude_amount(check_item + 1)
						duplct_pay_date(check_item)				= duplct_pay_date(check_item + 1)
						calculated_by_ytd(check_item) 			= calculated_by_ytd(check_item + 1)
						ytd_calc_notes(check_item) 				= ytd_calc_notes(check_item + 1)
						pay_detail_exists(check_item) 			= pay_detail_exists(check_item + 1)
						combined_into_one(check_item) 			= combined_into_one(check_item + 1)
						SNAP_dialog_display(check_item) 		= SNAP_dialog_display(check_item + 1)
						CASH_dialog_display(check_item) 		= CASH_dialog_display(check_item + 1)
					End If
				Next
			End If
			delete_check = UBound(pay_date) - 1
			call resize_check_list(delete_check)

		End If

	end sub

	public sub determine_bimonthly_dates()
        'allows for the selection of bimonthly dates if the pay frequency is 2 - Two Times Per Month since these are weird

        temp_array_of_pay_dates = pay_date

		Call sort_dates(temp_array_of_pay_dates)
		first_check = temp_array_of_pay_dates(0)
		last_check = temp_array_of_pay_dates(UBound(temp_array_of_pay_dates))

		list_of_days_of_checks = "~"
		the_day_of_month = ""
		third_paydate = FALSE
		If pay_freq = "2 - Two Times Per Month" AND (bimonthly_second = "" OR bimonthly_first = "") Then
			For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
				the_day_of_month = DatePart("d", pay_date(all_income))
				the_day_of_month = "~" & the_day_of_month & "~"
				If InStr(list_of_days_of_checks, the_day_of_month) = 0 Then
					the_day_of_month = replace(the_day_of_month, "~", "")
					list_of_days_of_checks = list_of_days_of_checks & the_day_of_month & "~"
				End If
			Next

			For each_day = 1 to 31
				each_day_spider = "~" & each_day & "~"
				If InStr(list_of_days_of_checks, each_day_spider) <> 0 Then
					If bimonthly_first = "" Then
						bimonthly_first = each_day
					ElseIf bimonthly_second = "" Then
						If each_day = 28 OR each_day = 29 OR each_day = 30 OR each_day = 31 Then
							bimonthly_second = "LAST"
						Else
							bimonthly_second = each_day
						End If
					Else
						third_paydate = TRUE
					End If
				End If
			Next

			If bimonthly_first = 28 OR bimonthly_first = 29 OR bimonthly_first = 30 OR bimonthly_first = 31 Then
				bimonthly_first = ""
				bimonthly_second = "LAST"
			End If

			If third_paydate = TRUE OR bimonthly_second = "" OR bimonthly_first = "" Then
				Call dialog_for_bimonthly_pay()
			End If

		End If
	end sub

	public sub determine_weekday_of_pay()
        'identify the weekday of pay for biweekly and weekly pay frequencies

        ReDim WEEKDAY_PAY_ARRAY(7)
		pd_by_wkdy = FALSE
		If actual_checks_provided = TRUE Then           'again, does not mater which way to budget is selected
			issues_with_frequency = FALSE               'default to false
			For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
				For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
					'conditional if it is the right panel AND the order matches - then do the thing you need to do
					If check_order(all_income) = order_number Then
						If NOT bonus_check(all_income) Then
							If pay_freq = "3 - Every Other Week" Then
								check_weekday = Weekday(pay_date(all_income))
								WEEKDAY_PAY_ARRAY(check_weekday) = WEEKDAY_PAY_ARRAY(check_weekday) + 1
								pd_by_wkdy = TRUE
							ElseIf pay_freq = "4 - Every Week" Then
								check_weekday = Weekday(pay_date(all_income))
								WEEKDAY_PAY_ARRAY(check_weekday) = WEEKDAY_PAY_ARRAY(check_weekday) + 1
								pd_by_wkdy = TRUE
							End If
						End If
					End If
				Next
			Next
		End If
		list_of_weekdays = "~"
		If pd_by_wkdy = TRUE Then
			two_paydays = FALSE
			For the_weekday = 1 to 7
				If WEEKDAY_PAY_ARRAY(the_weekday) <> 0 Then
					list_of_weekdays = list_of_weekdays & WeekDayName(the_weekday) & "~"
					If pay_weekday = "" or pay_weekday = WeekDayName(the_weekday) Then
						pay_weekday = WeekDayName(the_weekday)
						highest_count = WEEKDAY_PAY_ARRAY(the_weekday)
					ElseIf WEEKDAY_PAY_ARRAY(the_weekday) > highest_count Then
						pay_weekday = WeekDayName(the_weekday)
						highest_count = WEEKDAY_PAY_ARRAY(the_weekday)
						two_paydays = TRUE
					Else
						two_paydays = TRUE
					End If
				End If
			Next

			If two_paydays = TRUE Then
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 175, 85, "Weekday of Pay"
					DropListBox 95, 45, 75, 45, "Sunday"+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday", pay_weekday
					ButtonGroup ButtonPressed
						OkButton 120, 65, 50, 15
					Text 10, 10, 125, 10, "This job is paid weekly or biweekly."
					Text 10, 25, 165, 10, "Which day of the week is pay typically received?"
				EndDialog

				Dialog Dialog1
				save_your_work
			End If
		End If
	end sub

	public sub dialog_for_bimonthly_pay()
		'function to display and allow for change to the 2 days identified as the bimonthly pay dates.

        bimonthly_first = bimonthly_first & ""		'make the variables viewable in an EditBox
		bimonthly_second = bimonthly_second & ""
		If bimonthly_second = "LAST" Then		'aligning the internal process with the dialog information
			last_day_checkbox = checked
			bimonthly_second = ""
		End If

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 186, 90, "Days of Pay for Bimonthly"
			EditBox 55, 25, 25, 15, bimonthly_first
			EditBox 55, 45, 25, 15, bimonthly_second
			CheckBox 85, 50, 95, 10, "Second Day is LAST Day", last_day_checkbox
			ButtonGroup ButtonPressed
				OkButton 130, 70, 50, 15
			Text 5, 10, 150, 10, "Dates of Pay for BiMonthly Pay Frequency"
			Text 10, 30, 35, 10, "First Day"
			Text 10, 50, 45, 10, "Second Day"
		EndDialog

		Do
			the_err = ""

			dialog Dialog1						'show the dialog
			save_your_work

			bimonthly_first = trim(bimonthly_first)			'format the entries
			bimonthly_second = trim(bimonthly_second)

			'ensuring all information is entered and correct
			If bimonthly_first = "" Then the_err = the_err & vbNewLine & "* Enter the DAY of the month the first paycheck comes on."
			If IsNumeric(bimonthly_first) = False Then the_err = the_err & vbNewLine & "* The day for the first paycheck should be entered as a number, the day of the month first check is received."
			If bimonthly_second = "" Then
				If last_day_checkbox = unchecked Then the_err = the_err & vbNewLine & "* Enter the DAY of the month the second paycheck comes on. Or check the box indicating the second check falls on the last day of the month."
			Else
				If IsNumeric(bimonthly_second) = False Then the_err = the_err & vbNewLine & "* The day for the secon paycheck should be entered as a number, the day of the month first check is received. If this is always the last day of the month, check the box for the LAST day."
			End If
			If the_err <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & the_err
		Loop until the_err = ""

		'format the variables to be numbers
		bimonthly_first = bimonthly_first * 1
		If IsNumeric(bimonthly_second) = TRUE Then bimonthly_second = bimonthly_second * 1
		If last_day_checkbox = checked Then
			bimonthly_second = "LAST"
		ElseIf bimonthly_second = 30 OR bimonthly_second = 31 Then
			bimonthly_second = "LAST"
		End If
		save_your_work		'record this information
	end sub


	public sub evaluate_checks()
        'look at each check and create some formatting and details

        checks_exist = False
		For all_checks = 0 to UBound(pay_date)
			check_info_entered(all_checks) = False
			If pay_date(all_checks) <> "" Then
				checks_exist = True
				check_info_entered(all_checks) = True

                temp_string = ""
                excld_string = ""
                ' If pay_split_regular_amount(all_checks) <> "" Then
                If pay_split_bonus_amount(all_checks) <> "" Then temp_string = temp_string &  "$ " & pay_split_bonus_amount(all_checks) & " - Bonus "
                If pay_excld_bonus(all_checks) = checked Then
                    temp_string = temp_string & "(excluded) "
                    excld_string = excld_string & "Bonus Excluded "
                End If
                If pay_split_ot_amount(all_checks) <> "" Then temp_string = temp_string &  "$ " & pay_split_ot_amount(all_checks) & " - OT "
                If pay_excld_ot(all_checks) = checked Then
                    temp_string = temp_string & "(excluded) "
                    excld_string = excld_string & "OverTime Excluded"
                End If
                If pay_split_shift_diff_amount(all_checks) <> "" Then temp_string = temp_string &  "$ " & pay_split_shift_diff_amount(all_checks) & " - Shift Differential "
                If pay_excld_shift_diff(all_checks) = checked Then
                    temp_string = temp_string & "(excluded) "
                    excld_string = excld_string & "Shift Differential Excluded"
                End If
                If pay_split_tips_amount(all_checks) <> "" Then temp_string = temp_string &  "$ " & pay_split_tips_amount(all_checks) & " - Tips "
                If pay_excld_tips(all_checks) = checked Then
                    temp_string = temp_string & "(excluded) "
                    excld_string = excld_string & "Tips Excluded"
                End If
                If pay_split_other_amount(all_checks) <> "" Then temp_string = temp_string &  "$ " & pay_split_other_amount(all_checks) & " - " & pay_split_other_detail(all_checks) & " "
                If pay_excld_other(all_checks) = checked Then
                    temp_string = temp_string & "(excluded) "
                    excld_string = excld_string & pay_split_other_detail(all_checks) & " Excluded "
                End If
                split_check_string(all_checks) = temp_string
                split_check_excld_string(all_checks) = excld_string

				If exclude_entirely(all_checks) Then
					SNAP_info_string(all_checks) = "Check is excluded for SNAP. Reason: " & reason_to_exclude(all_checks)
					CASH_info_string(all_checks) = "Check is excluded for CASH. Reason: " & reason_to_exclude(all_checks)
				ElseIf bonus_check(all_checks) Then
					SNAP_info_string(all_checks) = "This is a Bonus Check and is excluded for SNAP."
					CASH_info_string(all_checks) = "This is a Bonus Check and is excluded for CASH."
				ElseIF IsNumeric(exclude_ALL_amount(all_checks)) Then
                    If exclude_ALL_amount(all_checks) <> 0 Then
                        SNAP_info_string(all_checks) = "$ " & exclude_ALL_amount(all_checks) & " of check is excluded for SNAP. Reason: " & reason_to_exclude(all_checks)
                        CASH_info_string(all_checks) = "$ " & exclude_ALL_amount(all_checks) & " of check is excluded for CASH. Reason: " & reason_to_exclude(all_checks)
                    End If
                End If


                If split_check_string(all_checks) <> "" Then
                    split_check_string(all_checks) = "Pay Details: $ " & pay_split_regular_amount(all_checks) & " - Reg. " & temp_string
                    SNAP_info_string(all_checks) = SNAP_info_string(all_checks) & split_check_string(all_checks)
                    CASH_info_string(all_checks) = CASH_info_string(all_checks) & split_check_string(all_checks)
                End If

                If exclude_from_SNAP(all_checks) = checked Then
                    SNAP_info_string(all_checks) = "Check is excluded for SNAP. Reason: " & reason_SNAP_amt_excluded(all_checks)
                ElseIf IsNumeric(exclude_SNAP_amount(all_checks)) Then
                    If exclude_SNAP_amount(all_checks) <> 0 Then SNAP_info_string(all_checks) = "$ " & exclude_SNAP_amount(all_checks) & " of check is excluded for SNAP. Reason: " & reason_SNAP_amt_excluded(all_checks)
                End If
                If SNAP_info_string(all_checks) = "" Then SNAP_info_string(all_checks) = "Entire check counted for SNAP."

                If exclude_from_CASH(all_checks) = checked Then
                    CASH_info_string(all_checks) = "Check is excluded for CASH. Reason: " & reason_CASH_amt_excluded(all_checks)
                ElseIf IsNumeric(exclude_CASH_amount(all_checks)) Then
                    If exclude_CASH_amount(all_checks) <> 0 Then CASH_info_string(all_checks) = "$ " & exclude_CASH_amount(all_checks) & " of check is excluded for CASH. Reason: " & reason_CASH_amt_excluded(all_checks)
                End If
                If CASH_info_string(all_checks) = "" Then CASH_info_string(all_checks) = "Entire check counted for CASH."

			End If
		Next

		estimate_exists = False
		If pay_per_hr <> "" and est_weekly_hrs <> "" Then estimate_exists = True

	end sub

	public sub evaluate_frequency()
        'make sure the frequency and checks align

		prev_date = ""              'setting some variables for the loop
		days_between_checks = ""
		'here we are going to see if there are checks out of line with the expected frequency.
		'These may be the correct paydates but later in the script we use the precise interval based on pay frequency to enter information
		If actual_checks_provided = TRUE Then           'again, does not mater which way to budget is selected
			issues_with_frequency = FALSE               'default to false
			For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
				For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
					'conditional if it is the right panel AND the order matches - then do the thing you need to do
					If check_order(all_income) = order_number Then
						'If first_check = "" Then first_check = pay_date(all_income)       'setting the first check to the panel if it has not been done
						list_of_dates = list_of_dates & vbNewLine & "Check Date: " & pay_date(all_income) & " Income: $" & gross_amount(all_income) & " Hours: " & hours(all_income)      'creating a readable list of the pay dates, amount, and hours
						view_pay_date(all_income) = pay_date(all_income)        'view pay date is the actual date that is always seen and typically is the same as the regular pay date
						frequency_issue(all_income) = FALSE                                           'defaulting this to false

						If NOT bonus_check(all_income) Then
							If prev_date <> "" Then     'we can't compare the first date to anything, so it skips the first date
								days_between_checks = DateDiff("d", prev_date, pay_date(all_income))      'determines how many days from one check to the next

								'if the number of days is more or less than exactly what we expect, we need clarification
								If pay_freq = "1 - One Time Per Month" Then
									If days_between_checks < 28 or days_between_checks > 31 Then
										issues_with_frequency = TRUE
										frequency_issue(all_income) = TRUE
										pay_date(all_income) = DateAdd("m", 1, prev_date)
									End If
								ElseIf pay_freq = "2 - Two Times Per Month" Then
									If bimonthly_second = "LAST" Then
										If DatePart("d", pay_date(all_income)) <> bimonthly_first Then
											day_after_pay = DateAdd("d", 1, pay_date(all_income))
											If DatePart("d", day_after_pay) <> 1 Then
												issues_with_frequency = TRUE
												frequency_issue(all_income) = TRUE
												month_to_use = DatePart("m", pay_date(all_income))
												year_to_use = DatePart("yyyy", pay_date(all_income))
												first_of_payMonth = month_to_use & "/1/" & year_to_use
												first_of_nextMonth = DateAdd("m", 1, first_of_payMonth)
												pay_date(all_income) = DateAdd("d", -1, first_of_nextMonth)
											End If
										End If

									Else
										If DatePart("d", pay_date(all_income)) <> bimonthly_first AND DatePart("d", pay_date(all_income)) <> bimonthly_second Then
											issues_with_frequency = TRUE
											frequency_issue(all_income) = TRUE
											month_to_use = DatePart("m", pay_date(all_income))
											year_to_use = DatePart("yyyy", pay_date(all_income))
											If DatePart("d", prev_date) = bimonthly_first Then pay_date(all_income) = month_to_use & "/" & bimonthly_second & "/" & year_to_use
											If DatePart("d", prev_date) = bimonthly_second Then pay_date(all_income) = month_to_use & "/" & bimonthly_first & "/" & year_to_use
										End If
									End If
								ElseIf pay_freq = "3 - Every Other Week" Then
									If duplct_pay_date(all_income) <> TRUE Then
										If WeekDayName(Weekday(pay_date(all_income))) <> pay_weekday OR days_between_checks <> 14 Then
											issues_with_frequency = TRUE
											frequency_issue(all_income) = TRUE
											pay_date(all_income) = DateAdd("d", 14, prev_date)
										End If
									End If
								ElseIf pay_freq = "4 - Every Week" Then
									If duplct_pay_date(all_income) <> TRUE Then
										If WeekDayName(Weekday(pay_date(all_income))) <> pay_weekday OR days_between_checks <> 7 Then
											issues_with_frequency = TRUE
											frequency_issue(all_income) = TRUE
											pay_date(all_income) = DateAdd("d", 7, prev_date)
										End If
									End If
								ElseIf pay_freq = "5 - Other" Then

								'REMOVE CODE
								Else        'this is code to determine the pay frequency for the worker but with all the other functionality - this is something the worker needs to provide
									If days_between_checks = 7 Then
										pay_freq = "4 - Every Week"
									ElseIf days_between_checks = 14 Then
										pay_freq = "3 - Every Other Week"
									ElseIf days_between_checks >= 14 AND days_between_checks <= 19 Then
										pay_freq = "2 - Two Times Per Month"
									ElseIf days_between_checks >= 28 AND days_between_checks <= 31 Then
										pay_freq = "1 - One Time Per Month"
									End If

								End If          'If pay_freq =
								prev_date = pay_date(all_income)
							Else
								If pay_freq = "3 - Every Other Week" OR pay_freq = "4 - Every Week" Then
									If WeekDayName(Weekday(pay_date(all_income))) <> pay_weekday Then
										issues_with_frequency = TRUE
										frequency_issue(all_income) = TRUE
										If pay_weekday = "Sunday" Then wkdy_nbr = 1
										If pay_weekday = "Monday" Then wkdy_nbr = 2
										If pay_weekday = "Tuesday" Then wkdy_nbr = 3
										If pay_weekday = "Wednesday" Then wkdy_nbr = 4
										If pay_weekday = "Thursday" Then wkdy_nbr = 5
										If pay_weekday = "Friday" Then wkdy_nbr = 6
										If pay_weekday = "Saturday" Then wkdy_nbr = 7
										date_difference = wkdy_nbr - Weekday(pay_date(all_income))
										pay_date(all_income) = DateAdd("d", date_difference, pay_date(all_income))

									Else
										prev_date = pay_date(all_income)      'saving this date as the one to compare to in the next loop
									End If
								Else
									prev_date = pay_date(all_income)      'saving this date as the one to compare to in the next loop
								End If
							End If          'If prev_date <> "" Then
						End If
					End If
				next
			next
			save_your_work

			If issues_with_frequency = TRUE Then        'if any checks did not align
				dlg_len = 85        'setting the base height

				For all_income = 0 to UBound(pay_date)       'increasing the height for each date with a frequency issue
					If frequency_issue(all_income) = TRUE Then dlg_len = dlg_len + 20
				Next

				'FREQUENCY ISSUE Dialog - the worker can update the view_pay_date to match if appropriate or they can confirm it is correct as is
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 251, dlg_len, "Review Pay Dates"
					Text 10, 10, 240, 10, "It appears one check does not fall in the expected pay schedule dates. "
					Text 10, 25, 230, 10, "This job is paid - " & pay_freq
					Text 10, 40, 65, 10, "Reported Pay Date"
					Text 85, 40, 75, 10, "Expected Pay Date"

					y_pos = 55
					For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
						For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
							'conditional if it is the right panel AND the order matches - then do the thing you need to do
							If check_order(all_income) = order_number Then
								If frequency_issue(all_income) = TRUE Then
									If view_pay_date(all_income) <> "" Then view_pay_date(all_income) = view_pay_date(all_income) & ""
									Text 10, y_pos, 10, 10, "**"
									EditBox 25, y_pos, 50, 15, view_pay_date(all_income)
									Text 95, y_pos + 5, 50, 10, pay_date(all_income)            'this cannot be changed here

									y_pos = y_pos + 20
								End If
							End If
						Next
					Next
					CheckBox 10, y_pos, 180, 10, "Check here if these pay dates are what was reported.", pay_dates_correct_checkbox
					ButtonGroup ButtonPressed
						OkButton 195, y_pos, 50, 15
				EndDialog

				Do
					Do
						dialog Dialog1      'showing the dialog
						save_your_work

						'worker must confirm the dates are correct
						If pay_dates_correct_checkbox = unchecked Then MsgBox "If the paydates reported in the paycheck information dates are incorrect. Correct them here. "
					Loop until pay_dates_correct_checkbox = checked
					call check_for_password(are_we_passworded_out)
				Loop until are_we_passworded_out = false

				For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
					If check_order(all_income) = order_number Then
						If frequency_issue(all_income) = TRUE Then
							'if there is more than a 6 day difference between the provided date and the expected date, we wil match the provided date
							If abs(DateDiff("d", view_pay_date(all_income), pay_date(all_income))) > 6 Then pay_date(all_income) = view_pay_date(all_income)
						End If
					End If
				Next
				save_your_work
			End If          'If issues_with_frequency = TRUE Then


			For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
				For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
					'conditional if it is the right panel AND the order matches - then do the thing you need to do
					If check_order(all_income) = order_number Then
						If first_check = "" Then first_check = pay_date(all_income)       'setting the first check to the panel if it has not been done
					End If
				Next
			Next
			default_start_date = first_check
			save_your_work
		End If          'If actual_checks_provided = TRUE Then
	end sub

	public sub evaluate_job_info()
        'identify what type of information was provided for the job

		word_for_freq = ""      'for displaying in the dialog
		If pay_freq = "1 - One Time Per Month" Then word_for_freq = "monthly"
		If pay_freq = "2 - Two Times Per Month" Then word_for_freq = "semi-monthly"
		If pay_freq = "3 - Every Other Week" Then word_for_freq = "biweekly"
		If pay_freq = "4 - Every Week" Then word_for_freq = "weekly"

		'setting for counting
		total_check_count = 0
		snap_check_count = 0
		cash_check_count = 0
		total_gross_amount = 0
		snap_budgeted_total = 0
		cash_budgeted_total = 0
		total_hours = 0
		snap_hours = 0
		cash_hours = 0

		actual_checks_provided = FALSE      'defaults for some logic coming up
		there_are_counted_checks = FALSE
		For all_checks = 0 to UBound(pay_date)
			If pay_date(all_checks) = "" Then check_info_entered(all_checks) = False
			If check_info_entered(all_checks) Then
				actual_checks_provided = TRUE
				total_check_count = total_check_count + 1
				total_gross_amount = total_gross_amount + gross_amount(all_checks)
				total_hours = total_hours + hours(all_checks)
				If NOT exclude_entirely(all_checks) Then
					there_are_counted_checks = TRUE
					If exclude_from_SNAP(all_checks) = unchecked Then
						snap_check_count = snap_check_count + 1
						If exclude_SNAP_amount(all_checks) = "" Then
							snap_budgeted_total = snap_budgeted_total + gross_amount(all_checks)
							snap_hours = snap_hours + hours(all_checks)
                        ElseIf IsNumeric(exclude_SNAP_amount(all_checks)) Then
							snap_budgeted_total = snap_budgeted_total + gross_amount(all_checks) - exclude_SNAP_amount(all_checks)
							snap_hours = snap_hours + hours(all_checks)
                            If IsNumeric(exclude_SNAP_hours(all_checks)) Then snap_hours = snap_hours - exclude_SNAP_hours(all_checks)
                        Else
                            snap_budgeted_total = snap_budgeted_total + gross_amount(all_checks)
                            snap_hours = snap_hours + hours(all_checks)
						End If
					End If
					If exclude_from_CASH(all_checks) = unchecked Then
						cash_check_count = cash_check_count + 1
						If exclude_CASH_amount(all_checks) = "" Then
							cash_budgeted_total = cash_budgeted_total + gross_amount(all_checks)
							cash_hours = cash_hours + hours(all_checks)
						ElseIf IsNumeric(exclude_CASH_amount(all_checks)) Then
							cash_budgeted_total = cash_budgeted_total + gross_amount(all_checks) - exclude_CASH_amount(all_checks)
							cash_hours = cash_hours + hours(all_checks)
                            If IsNumeric(exclude_CASH_hours(all_checks)) Then cash_hours = cash_hours - exclude_CASH_hours(all_checks)
                        Else
                            cash_budgeted_total = cash_budgeted_total + gross_amount(all_checks)
                            snacash_hoursp_hours = cash_hours + hours(all_checks)
						End If
					End If
				End If
			End If
		Next

		'If all anticiapated pay information has been provided, we look for a start date and define that anticiapated income is provided
		anticipated_income_provided = FALSE     'default
		If display_pay_per_hr <> "" AND display_hrs_per_wk <> "" AND pay_freq <> "" Then anticipated_income_provided = TRUE
		If reg_non_monthly <> "" AND numb_months <> "" Then anticipated_income_provided = TRUE

	end sub

	public sub find_missing_checks()
        'identify checks that are not provided

        missing_checks_list = ""
		expected_check_index = 0        'setting up for another loop to see if all the expected checks have in fact been provided.
		order_number = 1
		Do
			For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
				date_in_range = ""

				'conditional if it is the right panel AND the order matches - then do the thing you need to do
				If check_order(all_income) = order_number Then
					If NOT bonus_check(all_income) Then
						missing_check = FALSE       'defaulting this for each loop

						'here we are comparing each check from the ENTER PAY Dialog for this panel in order to the checks we expected to see
						'We can only get an accurate panel update if all the checks for the time frame provided are given - they can be excluded but they should be there
						'There are allowances here for some variation as sometimes paydates shift (ie holidays or extenuating circumstances)
						If pay_freq = "1 - One Time Per Month" Then
							date_in_range = DateDiff("d", pay_date(all_income), expected_check_array(expected_check_index))
							date_in_range = Abs(date_in_range)
							If date_in_range > 8 AND duplicate_pay_date(all_income) <> TRUE Then missing_check = TRUE      '8 day allowance
						ElseIf pay_freq = "2 - Two Times Per Month" Then
							date_in_range = DateDiff("d", pay_date(all_income), expected_check_array(expected_check_index))
							date_in_range = Abs(date_in_range)
							If date_in_range > 5 AND duplicate_pay_date(all_income) <> TRUE Then missing_check = TRUE      '5 day allowance
						ElseIf pay_freq = "3 - Every Other Week" Then
							date_in_range = DateDiff("d", pay_date(all_income), expected_check_array(expected_check_index))
							date_in_range = Abs(date_in_range)
							If date_in_range > 3 AND duplicate_pay_date(all_income) <> TRUE Then missing_check = TRUE      '3 day allowance
						ElseIf pay_freq = "4 - Every Week" Then
							date_in_range = DateDiff("d", pay_date(all_income), expected_check_array(expected_check_index))
							date_in_range = Abs(date_in_range)
							If date_in_range > 3 AND duplicate_pay_date(all_income) <> TRUE Then missing_check = TRUE      '3 day allowance
						End If

						If missing_check = TRUE Then        'if the date difference was too much then we save the date to a list
							missing_checks_list = missing_checks_list & "~" & expected_check_array(expected_check_index)
						End If
                        order_number = order_number + 1

                        If duplicate_pay_date(all_income) <> TRUE Then expected_check_index = expected_check_index + 1
						If order_number > order_ubound Then Exit For            'if we have reached the end of the entered checks OR the end of the expected checks, we need to leave the loop
						If expected_check_index > UBound(expected_check_array) Then Exit For
					Else
						order_number = order_number + 1
						If order_number > order_ubound Then Exit For            'if we have reached the end of the entered checks OR the end of the expected checks, we need to leave the loop
					End If

				End If
			Next
			If expected_check_index > UBound(expected_check_array) Then Exit Do     'if we have reached the end of the entered checks OR the end of the expected checks, we need to leave the loop
			' If order_number > order_ubound Then Exit Do
		Loop until order_number > order_ubound
	end sub

    public sub find_panel()
        'open the correct JOBS panel for the member and instance provided - allows for navigation if the panel is not found

        EMWriteScreen "JOBS", 20, 71
        transmit

        EMReadScreen JOBS_check, 4, 2, 45
        If JOBS_check <> "JOBS" Then Call Navigate_to_MAXIS_screen("STAT", "JOBS") 'navigate to JOBS for the right member and instance

        EMWriteScreen member, 20, 76
        EMWriteScreen instance, 20, 79
        transmit
        EMReadScreen confirm_same_employer, 30, 7, 42                  'double check the employer name because we don't want to have wrong income on the wrong panel and from month to month the instances may change
        the_new_instance = ""       'blanking this out
        If confirm_same_employer <> UCase(employer_with_underscores) Then      'if the name on the panel does not match the name in EARNED_INCOME_PANELS_ARRAY we have to figure this out
            'BUGGY CODE - this might be causing issues as there were a few reports but I cannot get it to confirm
            EMWriteScreen "JOBS", 20, 71        'go back to the first job for this person
            EMWriteScreen member, 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            try = 1         'we need an exit from the loop
            employers_read = ""
            Do
                EMReadScreen confirm_same_employer, 30, 7, 42      'now we read this on each panel
                employers_read = employers_read & "~~~~" & confirm_same_employer
                If confirm_same_employer = UCase(employer_with_underscores) Then               'if the panel has the employer name, then we set the new instance to EARNED_INCOME_PANELS_ARRAY
                    EMReadScreen the_new_instance, 1, 2, 73
                    instance = "0" & the_new_instance
                    Exit Do
                End If
                transmit                                'otherwise go to the next panel and repeat
                EMReadScreen last_jobs, 7, 24, 2
                try = try + 1
                If try = 15 Then Exit Do
            Loop until last_jobs = "ENTER A"            'This is when you can't transmit any more

            If the_new_instance = "" Then               'If they didn't matcj and we did not find it, this alerts the worker
                script_run_lowdown = script_run_lowdown & vbCr & "PANEL NOT FOUND In " & MAXIS_footer_month & "/" & MAXIS_footer_year
                temp_array = ""
                employers_read = trim(employers_read)
                temp_array = split(employers_read, "~~~~")
                for each job_read in temp_array
                    script_run_lowdown = script_run_lowdown & vbCr & "panel read - " & job_read
                Next

                Do
                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 196, 235, "Find the Correct Panel"
                        Text 10, 10, 155, 20, "The script has been unable to find the correct JOBS panel for the month " & MAXIS_footer_month & "/" & MAXIS_footer_year & "."
                        Text 10, 40, 185, 10, "The JOBS panel selected at the beginning of the script: "
                        Text 30, 50, 150, 10, employer_with_underscores
                        Text 10, 70, 165, 10, "The script read the following JOBS Employers:"
                        y_pos = 80
                        for each job_read in temp_array
                            Text 30, y_pos, 140, 10, job_read
                            y_pos = y_pos + 10
                        Next
                        Text 10, 140, 175, 10, "You can naviagate directly to the correct panel now. "
                        Text 15, 150, 175, 20, "Leave this dialog up and navigate in this MAXIS session to the panel for this job."
                        ButtonGroup ButtonPressed
                            PushButton 10, 180, 175, 15, "I have navigated to the Correct JOBS panel", panel_navigated_to_btn
                            PushButton 10, 205, 175, 15, "Skip the update of this job for the month " & MAXIS_footer_month & "/" & MAXIS_footer_year, skip_this_month_btn
                    EndDialog

                    dialog Dialog1

                Loop until ButtonPressed = panel_navigated_to_btn or ButtonPressed = skip_this_month_btn

                If ButtonPressed = skip_this_month_btn Then update_this_month = FALSE     'setting this to NOT update
                If ButtonPressed = panel_navigated_to_btn Then
                    old_instance = instance
                    EMReadScreen the_new_instance, 1, 2, 73

                    instance = "0" & the_new_instance

                    confirm_selection = MsgBox("The script will now update this panel:" & vbCr & "JOBS " & member & " " & instance & vbCr & vbCr & "Is this the panel you want updated with the income entered for the job - " & employer_with_underscores & "?",  vbSystemModal + vbExclamation + vbDefaultButton2 + VBYesNo, "CONFIRM PANEL UPDATE")
                    If confirm_selection = vbNo Then
                        instance = old_instance
                        update_this_month = FALSE
                    End If
                End If
                If ButtonPressed = skip_this_month_btn Then script_run_lowdown = script_run_lowdown & vbCr & "PRESSED BUTTON to Skip This Month"
                If ButtonPressed = panel_navigated_to_btn Then
                    script_run_lowdown = script_run_lowdown & vbCr & "PRESSED BUTTON to Panel navigated to"
                    If confirm_selection = vbNo Then script_run_lowdown = script_run_lowdown & vbCr & "MsgBox VBNo Pressed"
                    If confirm_selection = vbYes Then script_run_lowdown = script_run_lowdown & vbCr & "MsgBox VBYes Pressed"
                End If
            End If
        End If
    end sub

	public sub job_info_error_handling(err_msg)
        'dialog error handling for the job information provided

        'there are special criteria for using this verification code
		If verif_type = "? - EXPEDITED SNAP ONLY" Then
			apply_to_CASH 	= unchecked     					'only for SNAP
			apply_to_SNAP 	= checked
			apply_to_HC 	= unchecked
			apply_to_GRH 	= unchecked

			initial_month_mo = fs_appl_footer_month   		'month of application handling only
			initial_month_yr = fs_appl_footer_year
			If verif_explain = "" Then err_msg = err_msg & vbNewLine & "* If the verification code is '?' additional information about the verification needs to be added." 'need more explanation
			If all_pay_in_app_month = FALSE Then err_msg = err_msg & vbNewLine & "* Only income from the month of application should be entered when using '?' as this is only for income that is not sufficiently verified to be used to determine Expedited."    'this only happens if '?' is the verif code
		End If

		If apply_to_SNAP = unchecked AND apply_to_CASH = unchecked AND apply_to_HC = unchecked AND apply_to_GRH = unchecked Then err_msg = err_msg & vbNewLine & "* No programs have been selected that this income applies to. Chose at least one program that this income is budgeted for."
		verif_date = trim(verif_date)
		If verif_date = "" Then
			err_msg = err_msg & vbNewLine & "* Enter the date the pay information was received in the agency."
		ElseIf NOT IsDate(verif_date) Then
			err_msg = err_msg & vbNewLine & "* The date the pay information was received in the agency does not appear to be a valid date, review and update."
		ElseIF DateDiff("d", date, verif_date) > 0 Then
			err_msg = err_msg & vbNewLine & "* The date the pay information was received in the agency appears to be in the future, review and update."
		End If
		If pay_freq = "" Then err_msg = err_msg & vbNewLine & "* Select the pay frequency for this job."        'NEED to have a pay frequency

		If first_check <> "" Then
			end_of_month = initial_month_mo & "/1/" & initial_month_yr
			end_of_month = DateAdd("m", 1, end_of_month)
			end_of_month = DateAdd("d", -1, end_of_month)

			If DateDiff("d", first_check, end_of_month) < 0 Then err_msg = err_msg & vbNewLine & "* The check dates should start in or before the initial month to update. If no additional checks exist, change the initial month to update to the first month for which checks have been received."
		End If

		pay_per_hr = trim(pay_per_hr)       	'formatting
		hrs_per_wk = trim(hrs_per_wk)
		pay_freq = trim(pay_freq)

		If pay_per_hr <> "" and NOT IsNumeric(pay_per_hr) Then err_msg = err_msg & vbNewLine & "* The 'Rate of Pay/Hr' should be entered as number. Review and update the hourly pay rate."
		If hrs_per_wk <> "" and NOT IsNumeric(hrs_per_wk) Then err_msg = err_msg & vbNewLine & "* The 'Hours/wk' should be entered as a number. Review and update the number of weekly hours."

		reg_non_monthly = trim(reg_non_monthly)     'this is not currently in the dialog - FUTURE FUNCTIONALITY - need a lot of other handling to put this back in.
		numb_months = trim(numb_months)
		known_pay_date = trim(known_pay_date)

		If known_pay_date <> "" and NOT IsDate(known_pay_date) Then err_msg = err_msg & vbNewLine & "* The 'Known Pay Date' should be entered as a date. Review and update a known date in the Anticipated Income section."

		If anticipated_income_provided = TRUE and income_start_dt = "" Then err_msg = err_msg & vbNewLine & "* Enter an income start date, since anticipated pay dates cannot be determined without the initial pay date."

		If anticipated_income_provided = FALSE AND actual_checks_provided = FALSE Then          'there either needs to be checks OR anticipated income
			err_msg = err_msg & vbNewLine & "* Income information needs to be provided, either in the form of actual checks or anticipated income, hours, and rate of pay."
		End If
		If there_are_counted_checks = FALSE AND anticipated_income_provided = FALSE AND actual_checks_provided = TRUE Then
			If total_hours <> 0 Then
				pay_wage = total_gross_amount/total_hours
				pay_per_hr = pay_wage
				hrs_per_wk = 0
				anticipated_income_provided = TRUE
			Else
				err_msg = err_msg & vbNewLine & "* All the checks listed are excluded and no anticipated income estimate is provided. In order to udate a case and budget income there needs to be counted income."
			End If
		End If
		If known_pay_date <> "" AND IsDate(known_pay_date) = FALSE Then err_msg = err_msg & vbNewLine & "* A known pay date needs to be entered as a date. Check the entry."

	end sub

	public sub make_known_date_earlier()
        'the known date needs to start at the beginning so we sometimes need to calculate back

		If known_pay_date <> "" Then
			known_pay_date = DateValue(known_pay_date)
			the_initial_month = DateValue(initial_month_mo & "/1/" & initial_month_yr)
			If DateDiff("d", known_pay_date, the_initial_month) < 0 Then
				the_month_before = DateAdd("m", -1, the_initial_month)

				Do
					If pay_freq = "1 - One Time Per Month" Then       'each next date is determined by the pay frequency
						the_previous_pay = DateAdd("m", -1, known_pay_date)
					ElseIf pay_freq = "2 - Two Times Per Month" Then
						If DatePart("d", known_pay_date) = bimonthly_first Then         'If we are at the first check of the month, we need to go to the second
							If bimonthly_second = "LAST" Then
								pay_month = DatePart("m", known_pay_date)
								pay_year = DatePart("yyyy", known_pay_date)
								this_month = DateValue(pay_month & "/1/" & pay_year)
								the_previous_pay = DateAdd("d", -1, this_month)
							Else
								next_pay = DateAdd("m", -1, known_pay_date)                                                            'go to the next month
								next_pay_month = DatePart("m", next_pay)
								next_pay_year = DatePart("yyyy", next_pay)

								the_previous_pay = next_pay_month & "/" & bimonthly_second & "/" & next_pay_year
							End If
						Else
							next_pay_month = DatePart("m", known_pay_date)
							next_pay_year = DatePart("yyyy", known_pay_date)
							the_previous_pay = next_pay_month & "/" & bimonthly_first & "/" & next_pay_year
						End If
					ElseIf pay_freq = "3 - Every Other Week" Then
						the_previous_pay = DateAdd("d", -14, known_pay_date)
					ElseIf pay_freq = "4 - Every Week" Then
						the_previous_pay = DateAdd("d", -7, known_pay_date)
					End If

					known_pay_date = the_previous_pay
				Loop Until DateDiff("d", known_pay_date, the_initial_month) >= 0
			End If
			the_initial_month = ""
		End If
	end sub

	public sub order_checks()
        'put the checks in date order

		temp_array_of_pay_dates = pay_date

		Call sort_dates(temp_array_of_pay_dates)
		first_check = temp_array_of_pay_dates(0)
		last_check = temp_array_of_pay_dates(UBound(temp_array_of_pay_dates))
		all_pay_in_app_month = False
        If IsDate(fs_appl_date) Then all_pay_in_app_month = TRUE

		the_counter = 1
		assessed_checks_list = "~"
		For each check in temp_array_of_pay_dates
			For all_income = 0 to UBound(pay_date)           'Now loop through all of the listed income - again
				If check_info_entered(all_income) Then
					If check = pay_date(all_income) Then
						pay_date(all_income) = DateValue(pay_date(all_income))
						check_order(all_income) = the_counter
						If InStr(assessed_checks_list, "~" & pay_date(all_income) & "~") <> 0 Then
							duplicate_pay_date(all_income) = True
						Else
							assessed_checks_list = assessed_checks_list & pay_date(all_income) & "~"
						End If
						order_ubound = the_counter
						the_counter = the_counter + 1
					End If
					'this is a little messy - BUGGY CODE - works fine but maybe needs a logic upgrade to be more elegant
                    If IsDate(fs_appl_date) Then
                        If NOT (DatePart("m", fs_appl_date) = DatePart("m", pay_date(all_income)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", pay_date(all_income))) Then
                            all_pay_in_app_month = FALSE
                        End If
                    End If
				End If
			Next
		Next
	end sub

	public sub read_panel()
        'get current panel information
		ReDim pay_date(0)
		ReDim gross_amount(0)
		ReDim hours(0)
		ReDim exclude_entirely(0)
		ReDim exclude_from_SNAP(0)
		ReDim exclude_from_CASH(0)
		ReDim reason_to_exclude(0)
		ReDim exclude_ALL_amount(0)
        ReDim exclude_ALL_hours(0)
		ReDim exclude_SNAP_amount(0)
        ReDim exclude_SNAP_hours(0)
		ReDim exclude_CASH_amount(0)
        ReDim exclude_CASH_hours(0)
		ReDim SNAP_info_string(0)
		ReDim CASH_info_string(0)
		ReDim check_order(0)
		ReDim view_pay_date(0)
		ReDim frequency_issue(0)
		ReDim future_check(0)
		ReDim duplicate_pay_date(0)
		ReDim reason_SNAP_amt_excluded(0)
		ReDim reason_CASH_amt_excluded(0)
		ReDim pay_detail_btn(0)
		ReDim check_info_entered(0)
		ReDim bonus_check(0)
		ReDim pay_split_regular_amount(0)
		ReDim pay_split_bonus_amount(0)
		ReDim pay_split_ot_amount(0)
        ReDim pay_split_ot_hours(0)
		ReDim pay_split_shift_diff_amount(0)
		ReDim pay_split_tips_amount(0)
		ReDim pay_split_other_amount(0)
		ReDim pay_split_other_detail(0)
		ReDim pay_excld_bonus(0)
		ReDim pay_excld_ot(0)
		ReDim pay_excld_shift_diff(0)
		ReDim pay_excld_tips(0)
		ReDim pay_excld_other(0)
		ReDim split_check_string(0)
        ReDim split_check_excld_string(0)
        ReDim split_exclude_amount(0)
		ReDim duplct_pay_date(0)
		ReDim calculated_by_ytd(0)
		ReDim ytd_calc_notes(0)
		ReDim pay_detail_exists(0)
		ReDim combined_into_one(0)
		ReDim SNAP_dialog_display(0)
		ReDim CASH_dialog_display(0)

        cash_array_info_exists = False
        ReDim cash_info_cash_mo_yr(0)
        ReDim cash_info_retro_mo_yr(0)
        ReDim cash_info_retro_updtd(0)
        ReDim cash_info_prosp_updtd(0)
        ReDim cash_info_mo_retro_pay(0)
        ReDim cash_info_mo_retro_hrs(0)
        ReDim cash_info_mo_prosp_pay(0)
        ReDim cash_info_mo_prosp_hrs(0)

		next_check_btn = 150
		cancel_check_btn = 160
		save_details_btn = 170
		delete_check_btn = 180

		checks_exist = False
		estimate_exists = False

		'Reading the information from the panel
		'FUTURE FUNCTIONALITY - add ability to read current income from the panel/PIC etc. so that partial work can be screen scraped instead of having to retype it
		EMReadScreen type_of_job, 1, 5, 34
		EMReadScreen job_verif, 25, 6, 34
		EMReadScreen listed_hrly_wage, 6, 6, 75
		EMReadScreen employer_name, 30, 7, 42
		EMReadScreen start_date, 8, 9, 35
		EMReadScreen end_date, 8, 9, 49
		EMReadScreen frequency, 1, 18, 35
		EMReadScreen current_verif, 27, 6, 34
		EMReadScreen updated_date, 8, 21, 55

		If type_of_job = "J" Then job_type = "J - WIOA"       'setting the full detail to the array instead of a single letter code
		If type_of_job = "W" Then job_type = "W - Wages"
		If type_of_job = "E" Then job_type = "E - EITC"
		If type_of_job = "G" Then job_type = "G - Experience Works"
		If type_of_job = "F" Then job_type = "F - Federal Work Study"
		If type_of_job = "S" Then job_type = "S - State Work Study"
		If type_of_job = "O" Then job_type = "O - Other"
		If type_of_job = "C" Then job_type = "C - Contract Income"
		If type_of_job = "T" Then job_type = "T - Training Program"
		If type_of_job = "P" Then job_type = "P - Service Program"
		If type_of_job = "R" Then job_type = "R - Rehab Program"
		If type_of_job = "N" Then job_type = "N - Census Income"

		'formatting the information from the panel and adding it to the EARNED_INCOME_PANELS_ARRAY
		verif_type = trim(job_verif)
		employer = replace(employer_name, "_", "")
		employer_with_underscores = employer_name
		hourly_wage = trim(listed_hrly_wage)
		income_start_dt = replace(start_date, " ", "/")
		income_end_dt = replace(end_date, " ", "/")
		updated_date = replace(updated_date, " ", "/")
		If income_start_dt = "__/__/__" Then income_start_dt = ""
		If income_end_dt = "__/__/__" Then income_end_dt = ""
		old_verif = trim(current_verif)
		' EARNED_INCOME_PANELS_ARRAY(income_list_indct, the_panel) = "NONE"       'This is where all of the array items from LIST_OF_INCOME_ARRAY will be added that are associated with this panel
		new_panel = FALSE      'identifies if a panel was created by the script or not - these are currently existing - changes CNote
		EMReadScreen first_retro_check, 8, 12, 25   'first check on retro side
		If first_retro_check = "__ __ __" Then
			first_retro_check = ""
		Else
			first_retro_check = replace(first_retro_check, " ", "/")
			first_retro_check = DateValue(first_retro_check)
		End If
	end sub

	public sub resize_check_list(top_index)
        'Update the check size of the arrays to match the number of checks entered
		ReDim preserve pay_date(top_index)
		ReDim preserve gross_amount(top_index)
		ReDim preserve hours(top_index)
		ReDim preserve exclude_entirely(top_index)
		ReDim preserve exclude_from_SNAP(top_index)
		ReDim preserve exclude_from_CASH(top_index)
		ReDim preserve reason_to_exclude(top_index)
		ReDim preserve exclude_ALL_amount(top_index)
        ReDim preserve exclude_ALL_hours(top_index)
		ReDim preserve exclude_SNAP_amount(top_index)
        ReDim preserve exclude_SNAP_hours(top_index)
		ReDim preserve exclude_CASH_amount(top_index)
        ReDim preserve exclude_CASH_hours(top_index)
		ReDim preserve SNAP_info_string(top_index)
		ReDim preserve CASH_info_string(top_index)
		ReDim preserve check_order(top_index)
		ReDim preserve view_pay_date(top_index)
		ReDim preserve frequency_issue(top_index)
		ReDim preserve future_check(top_index)
		ReDim preserve duplicate_pay_date(top_index)
		ReDim preserve reason_SNAP_amt_excluded(top_index)
		ReDim preserve reason_CASH_amt_excluded(top_index)
		ReDim preserve pay_detail_btn(top_index)
		ReDim preserve check_info_entered(top_index)
		ReDim preserve bonus_check(top_index)
		ReDim preserve pay_split_regular_amount(top_index)
		ReDim preserve pay_split_bonus_amount(top_index)
		ReDim preserve pay_split_ot_amount(top_index)
        ReDim preserve pay_split_ot_hours(top_index)
		ReDim preserve pay_split_shift_diff_amount(top_index)
		ReDim preserve pay_split_tips_amount(top_index)
		ReDim preserve pay_split_other_amount(top_index)
		ReDim preserve pay_split_other_detail(top_index)
		ReDim preserve pay_excld_bonus(top_index)
		ReDim preserve pay_excld_ot(top_index)
		ReDim preserve pay_excld_shift_diff(top_index)
		ReDim preserve pay_excld_tips(top_index)
		ReDim preserve pay_excld_other(top_index)
		ReDim preserve split_check_string(top_index)
        ReDim preserve split_check_excld_string(top_index)
        ReDim preserve split_exclude_amount(top_index)
		ReDim preserve duplct_pay_date(top_index)
		ReDim preserve calculated_by_ytd(top_index)
		ReDim preserve ytd_calc_notes(top_index)
		ReDim preserve pay_detail_exists(top_index)
		ReDim preserve combined_into_one(top_index)
		ReDim preserve SNAP_dialog_display(top_index)
		ReDim preserve CASH_dialog_display(top_index)
	end sub

    public sub restore_info(header, info)
        'getting information from the txt file for save_your_work functionality

        If header = "new_panel"                         Then Call read_txt_value(new_panel,                 info, "Boolean")
        If header = "member"                            Then Call read_txt_value(member,                    info, "String")
        If header = "instance"                          Then Call read_txt_value(instance,                  info, "String")
        If header = "employer"                          Then Call read_txt_value(employer,                  info, "String")
        If header = "employer_with_underscores"         Then Call read_txt_value(employer_with_underscores, info, "String")
        If header = "job_type"                          Then Call read_txt_value(job_type,                  info, "String")
        If header = "verif_type"                        Then Call read_txt_value(verif_type,                info, "String")
        If header = "hourly_wage"                       Then Call read_txt_value(hourly_wage,               info, "Amount")
        If header = "snap_hourly_wage"                  Then Call read_txt_value(snap_hourly_wage,          info, "Amount")
        If header = "cash_hourly_wage"                  Then Call read_txt_value(cash_hourly_wage,          info, "Amount")
        If header = "income_start_dt"                   Then Call read_txt_value(income_start_dt,           info, "String")
        If header = "income_end_dt"                     Then Call read_txt_value(income_end_dt,             info, "Date")
        If header = "pay_freq"                          Then Call read_txt_value(pay_freq,                  info, "String")
        If header = "updated_date"                      Then Call read_txt_value(updated_date,              info, "Date")
        If header = "old_verif"                         Then Call read_txt_value(old_verif,                 info, "String")
        If header = "first_retro_check"                 Then Call read_txt_value(first_retro_check,         info, "Date")
        If header = "initial_month_mo"                  Then Call read_txt_value(initial_month_mo,          info, "String")
        If header = "initial_month_yr"                  Then Call read_txt_value(initial_month_yr,          info, "String")
        If header = "hrs_per_wk"                        Then Call read_txt_value(hrs_per_wk,                info, "Amount")
        If header = "pay_per_hr"                        Then Call read_txt_value(pay_per_hr,                info, "Amount")
        If header = "display_hrs_per_wk"                Then Call read_txt_value(display_hrs_per_wk,        info, "String")
        If header = "display_pay_per_hr"                Then Call read_txt_value(display_pay_per_hr,        info, "String")
        If header = "known_pay_date"                    Then Call read_txt_value(known_pay_date,            info, "Date")
        If header = "word_for_freq"                     Then Call read_txt_value(word_for_freq,             info, "String")
        If header = "apply_to_SNAP"                     Then Call read_txt_value(apply_to_SNAP,             info, "Checkbox")
        If header = "apply_to_CASH"                     Then Call read_txt_value(apply_to_CASH,             info, "Checkbox")
        If header = "apply_to_HC"                       Then Call read_txt_value(apply_to_HC,               info, "Checkbox")
        If header = "apply_to_GRH"                      Then Call read_txt_value(apply_to_GRH,              info, "Checkbox")
        If header = "prog_list"                         Then Call read_txt_value(prog_list,                 info, "String")
        If header = "verif_date"                        Then Call read_txt_value(verif_date,                info, "String")
        If header = "verif_explain"                     Then Call read_txt_value(verif_explain,             info, "String")
        If header = "selection_rsn"                     Then Call read_txt_value(selection_rsn,             info, "String")
        If header = "income_excluded_cash_reason"       Then Call read_txt_value(income_excluded_cash_reason, info, "String")
        If header = "hc_budg_notes"                     Then Call read_txt_value(hc_budg_notes,             info, "String")
        If header = "hc_retro"                          Then Call read_txt_value(hc_retro,                  info, "Boolean")
        If header = "EI_panel_vbYes"                    Then Call read_txt_value(EI_panel_vbYes,            info, "Boolean")
        If header = "spoke_with"                        Then Call read_txt_value(spoke_with,                info, "String")
        If header = "convo_detail"                      Then Call read_txt_value(convo_detail,              info, "String")
        If header = "paycheck_list_title"               Then Call read_txt_value(paycheck_list_title,       info, "String")
        If header = "excl_cash_rsn"                     Then Call read_txt_value(excl_cash_rsn,             info, "String")
        If header = "budget_confirmed"                  Then Call read_txt_value(budget_confirmed,          info, "Boolean")
        If header = "SNAP_list_of_excluded_pay_dates"   Then Call read_txt_value(SNAP_list_of_excluded_pay_dates,   info, "String")
        If header = "CASH_list_of_excluded_pay_dates"   Then Call read_txt_value(CASH_list_of_excluded_pay_dates,   info, "String")
        If header = "pay_weekday"                       Then Call read_txt_value(pay_weekday,               info, "String")
        If header = "income_received"                   Then Call read_txt_value(income_received,           info, "Boolean")
        If header = "all_pay_in_app_month"              Then Call read_txt_value(all_pay_in_app_month,      info, "Boolean")
        If header = "there_are_counted_checks"          Then Call read_txt_value(there_are_counted_checks,  info, "Boolean")
        If header = "actual_checks_provided"            Then Call read_txt_value(actual_checks_provided,    info, "Boolean")
        If header = "anticipated_income_provided"       Then Call read_txt_value(anticipated_income_provided,info, "Boolean")
        If header = "missing_checks_list"               Then Call read_txt_value(missing_checks_list,       info, "String")
        If header = "issues_with_frequency"             Then Call read_txt_value(issues_with_frequency,     info, "Boolean")
        If header = "default_start_date"                Then Call read_txt_value(default_start_date,        info, "Date")
        If header = "bimonthly_first"                   Then Call read_txt_value(bimonthly_first,           info, "String")
        If header = "bimonthly_second"                  Then Call read_txt_value(bimonthly_second,          info, "String")
        If header = "ignore_antic"                      Then Call read_txt_value(ignore_antic,              info, "Boolean")
        If header = "pick_one"                          Then Call read_txt_value(pick_one,                  info, "Number")
        If header = "days_to_add"                       Then Call read_txt_value(days_to_add,               info, "Number")
        If header = "months_to_add"                     Then Call read_txt_value(months_to_add,             info, "Number")
        If header = "first_check"                       Then Call read_txt_value(first_check,               info, "Date")
        If header = "last_check"                        Then Call read_txt_value(last_check,                info, "Date")
        If header = "order_ubound"                      Then Call read_txt_value(order_ubound,              info, "Number")
        If header = "checks_exist"                      Then Call read_txt_value(checks_exist,              info, "Boolean")
        If header = "estimate_exists"                   Then Call read_txt_value(estimate_exists,           info, "Boolean")
        If header = "gross_max_string_len"              Then Call read_txt_value(gross_max_string_len,      info, "Number")
        If header = "total_check_count"                 Then Call read_txt_value(total_check_count,         info, "Number")
        If header = "total_gross_amount"                Then Call read_txt_value(total_gross_amount,        info, "Amount")
        If header = "total_hours"                       Then Call read_txt_value(total_hours,               info, "Amount")
        If header = "ave_hrs_per_pay"                   Then Call read_txt_value(ave_hrs_per_pay,           info, "Amount")
        If header = "ave_inc_per_pay"                   Then Call read_txt_value(ave_inc_per_pay,           info, "Amount")
        If header = "monthly_income"                    Then Call read_txt_value(monthly_income,            info, "Amount")
        If header = "snap_check_count"                  Then Call read_txt_value(snap_check_count,          info, "Number")
        If header = "snap_budgeted_total"               Then Call read_txt_value(snap_budgeted_total,       info, "Amount")
        If header = "snap_hours"                        Then Call read_txt_value(snap_hours,                info, "Amount")
        If header = "snap_ave_hrs_per_pay"              Then Call read_txt_value(snap_ave_hrs_per_pay,      info, "Amount")
        If header = "snap_ave_inc_per_pay"              Then Call read_txt_value(snap_ave_inc_per_pay,      info, "Amount")
        If header = "snap_hrs_per_wk"                   Then Call read_txt_value(snap_hrs_per_wk,           info, "Amount")
        If header = "SNAP_monthly_income"               Then Call read_txt_value(SNAP_monthly_income,       info, "Amount")
        If header = "cash_check_count"                  Then Call read_txt_value(cash_check_count,          info, "Number")
        If header = "cash_budgeted_total"               Then Call read_txt_value(cash_budgeted_total,       info, "Amount")
        If header = "cash_hours"                        Then Call read_txt_value(cash_hours,                info, "Amount")
        If header = "cash_ave_hrs_per_pay"              Then Call read_txt_value(cash_ave_hrs_per_pay,      info, "Amount")
        If header = "cash_ave_inc_per_pay"              Then Call read_txt_value(cash_ave_inc_per_pay,      info, "Amount")
        If header = "cash_hrs_per_wk"                   Then Call read_txt_value(cash_hrs_per_wk,           info, "Amount")
        If header = "CASH_monthly_income"               Then Call read_txt_value(CASH_monthly_income,       info, "Amount")
        If header = "months_updated"                    Then Call read_txt_value(months_updated,            info, "String")
        If header = "income_lumped_mo"                  Then Call read_txt_value(income_lumped_mo,          info, "String")
        If header = "lump_reason"                       Then Call read_txt_value(lump_reason,               info, "String")
        If header = "act_checks_lumped"                 Then Call read_txt_value(act_checks_lumped,         info, "String")
        If header = "est_checks_lumped"                 Then Call read_txt_value(est_checks_lumped,         info, "String")
        If header = "lump_gross"                        Then Call read_txt_value(lump_gross,                info, "Amount")
        If header = "lump_hrs"                          Then Call read_txt_value(lump_hrs,                  info, "Amount")
        If header = "mo_w_more_5_chcks"                 Then Call read_txt_value(mo_w_more_5_chcks,         info, "String")
        If header = "update_future"                     Then Call read_txt_value(update_future,             info, "Checkbox")
        If header = "cash_array_info_exists"            Then Call read_txt_value(cash_array_info_exists,    info, "Boolean")
        ' If header = "numb_months"                       Then Call read_txt_value(numb_months,               info, "")         'Currently never assigned
        ' If header = "reg_non_monthly"                   Then Call read_txt_value(reg_non_monthly,           info, "")         'Currently never assigned
        ' If header = "update_this_month"                 Then Call read_txt_value(update_this_month,         info, "")         'Need to leave this as a default
        ' If header = "updates_to_display"                Then Call read_txt_value(updates_to_display,        info, "")         'Need to leave this as a default

        ' If header = "snap_anticipated_pay_array"        Then Call read_txt_value(snap_anticipated_pay_array,info, "")
        ' If header = "cash_anticipated_pay_array"        Then Call read_txt_value(cash_anticipated_pay_array,info, "")

        If header = "cash_anticipated_pay_array" Then Call read_txt_array(cash_anticipated_pay_array, info, "String", "|", True)
        '     temp_array = split(info, "|")
        '     ReDim cash_anticipated_pay_array(UBound(temp_array))
        '     For the_thing = 0 to UBound(temp_array)
        '         cash_anticipated_pay_array(the_thing) = temp_array(the_thing)
        '     Next
        ' End If

        If header = "snap_anticipated_pay_array" Then Call read_txt_array(snap_anticipated_pay_array, info, "String", "|", True)
        '     temp_array = split(info, "|")
        '     ReDim snap_anticipated_pay_array(UBound(temp_array))
        '     For the_thing = 0 to UBound(temp_array)
        '         snap_anticipated_pay_array(the_thing) = temp_array(the_thing)
        '     Next
        ' End If

        If header = "expected_check_array" Then Call read_txt_array(expected_check_array, info, "Date", "|", True)
        '     temp_array = split(info, "|")
        '     ReDim expected_check_array(UBound(temp_array))
        '     For the_thing = 0 to UBound(temp_array)
        '         expected_check_array(the_thing) = temp_array(the_thing)
        '     Next
        ' End If

        If header = "this_month_checks_array" Then Call read_txt_array(this_month_checks_array, info, "Date", "|", True)
        '     temp_array = split(info, "|")
        '     ReDim this_month_checks_array(UBound(temp_array))
        '     For the_thing = 0 to UBound(temp_array)
        '         this_month_checks_array(the_thing) = temp_array(the_thing)
        '     Next
        ' End If

        If header = "retro_month_checks_array" Then Call read_txt_array(retro_month_checks_array, info, "Date", "|", True)
        '     temp_array = split(info, "|")
        '     ReDim retro_month_checks_array(UBound(temp_array))
        '     For the_thing = 0 to UBound(temp_array)
        '         retro_month_checks_array(the_thing) = temp_array(the_thing)
        '     Next
        ' End If

        If header = "CASH_ARRAY" Then
            array_ubound = info
            If NOT IsNumeric(array_ubound) Then array_ubound = 0
            array_ubound = array_ubound * 1

            ReDim cash_info_cash_mo_yr(array_ubound)
            ReDim cash_info_retro_mo_yr(array_ubound)
            ReDim cash_info_retro_updtd(array_ubound)
            ReDim cash_info_prosp_updtd(array_ubound)
            ReDim cash_info_mo_retro_pay(array_ubound)
            ReDim cash_info_mo_retro_hrs(array_ubound)
            ReDim cash_info_mo_prosp_pay(array_ubound)
            ReDim cash_info_mo_prosp_hrs(array_ubound)
        End If

        If header = "PAYCHECK_ARRAY" Then
            array_ubound = info
            If NOT IsNumeric(array_ubound) Then array_ubound = 0
            array_ubound = array_ubound * 1
            resize_check_list(array_ubound)
        End If

        If header = "cash_info_cash_mo_yr" Then Call read_txt_array(cash_info_cash_mo_yr, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         cash_info_cash_mo_yr(the_thing) = temp_array(the_thing)
        '     next
        ' End If
        If header = "cash_info_retro_mo_yr" Then Call read_txt_array(cash_info_retro_mo_yr, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         cash_info_retro_mo_yr(the_thing) = temp_array(the_thing)
        '     next
        ' End If
        If header = "cash_info_retro_updtd" Then Call read_txt_array(cash_info_retro_updtd, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         If UCASE(temp_array(the_thing)) = "FALSE" Then cash_info_retro_updtd(the_thing) = False
        '         If UCASE(temp_array(the_thing)) = "TRUE" Then cash_info_retro_updtd(the_thing) = True
        '     next
        ' End If
        If header = "cash_info_prosp_updtd" Then Call read_txt_array(cash_info_prosp_updtd, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         If UCASE(temp_array(the_thing)) = "FALSE" Then cash_info_prosp_updtd(the_thing) = False
        '         If UCASE(temp_array(the_thing)) = "TRUE" Then cash_info_prosp_updtd(the_thing) = True
        '     next
        ' End If
        If header = "cash_info_mo_retro_pay" Then Call read_txt_array(cash_info_mo_retro_pay, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         cash_info_mo_retro_pay(the_thing) = temp_array(the_thing)
        '         If NOT IsNumeric(cash_info_mo_retro_pay(the_thing)) Then temp_array(the_thing) = 0
        '         cash_info_mo_retro_pay(the_thing) = FormatNumber(cash_info_mo_retro_pay(the_thing), 2,,0)
        '     next
        ' End If
        If header = "cash_info_mo_retro_hrs" Then Call read_txt_array(cash_info_mo_retro_hrs, info, "Integer", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         cash_info_mo_retro_hrs(the_thing) = temp_array(the_thing)
        '         If NOT IsNumeric(cash_info_mo_retro_hrs(the_thing)) Then cash_info_mo_retro_hrs(the_thing) = 0
        '         cash_info_mo_retro_hrs(the_thing) = Round(cash_info_mo_retro_hrs(the_thing))
        '     next
        ' End If
        If header = "cash_info_mo_prosp_pay" Then Call read_txt_array(cash_info_mo_prosp_pay, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         cash_info_mo_prosp_pay(the_thing) = temp_array(the_thing)
        '         If NOT IsNumeric(cash_info_mo_prosp_pay(the_thing)) Then cash_info_mo_prosp_pay(the_thing) = 0
        '         cash_info_mo_prosp_pay(the_thing) = FormatNumber(cash_info_mo_prosp_pay(the_thing), 2,,0)
        '     next
        ' End If
        If header = "cash_info_mo_prosp_hrs" Then Call read_txt_array(cash_info_mo_prosp_hrs, info, "Integer", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for the_thing = 0 to UBound(temp_array)
        '         cash_info_mo_prosp_hrs(the_thing) = temp_array(the_thing)
        '         If NOT IsNumeric(cash_info_mo_prosp_hrs(the_thing)) Then cash_info_mo_prosp_hrs(the_thing) = 0
        '         cash_info_mo_prosp_hrs(the_thing) = Round(cash_info_mo_prosp_hrs(the_thing))
        '     next
        ' End If


        If header = "pay_date" Then Call read_txt_array(pay_date, info, "Date", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_date(cow) = temp_array(cow)
        '         If IsDate(pay_date(cow)) Then pay_date(cow) = DateAdd("d", 0, pay_date(cow))
        '     next
        ' End If
        If header = "gross_amount" Then Call read_txt_array(gross_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         gross_amount(cow) = temp_array(cow)
        '         If gross_amount(cow) Then gross_amount(cow) = gross_amount(cow)
        '         If IsNumeric(gross_amount(cow)) Then gross_amount(cow) = FormatNumber(gross_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "hours" Then Call read_txt_array(hours, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         hours(cow) = temp_array(cow)
        '         If hours(cow) Then hours(cow) = hours(cow)
        '         If IsNumeric(hours(cow)) Then hours(cow) = FormatNumber(hours(the_thing), 2,,0)
        '     next
        ' End If
        If header = "exclude_entirely" Then Call read_txt_array(exclude_entirely, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_entirely(cow) = temp_array(cow)
        '         If UCase(exclude_entirely(cow)) = "TRUE" Then exclude_entirely(cow) = True
        '         If UCase(exclude_entirely(cow)) = "FALSE" Then exclude_entirely(cow) = False
        '     next
        ' End If
        If header = "exclude_from_SNAP" Then Call read_txt_array(exclude_from_SNAP, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_from_SNAP(cow) = temp_array(cow)
        '         If exclude_from_SNAP(cow) = "" Then exclude_from_SNAP(cow) = unchecked
        '         exclude_from_SNAP(cow) = exclude_from_SNAP(cow) * 1
        '     next
        ' End If
        If header = "exclude_from_CASH" Then Call read_txt_array(exclude_from_CASH, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_from_CASH(cow) = temp_array(cow)
        '         If exclude_from_CASH(cow) = "" Then exclude_from_CASH(cow) = unchecked
        '         exclude_from_CASH(cow) = exclude_from_CASH(cow) * 1
        '     next
        ' End If
        If header = "reason_to_exclude" Then Call read_txt_array(reason_to_exclude, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         reason_to_exclude(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "exclude_ALL_amount" Then Call read_txt_array(exclude_ALL_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_ALL_amount(cow) = temp_array(cow)
        '         If IsNumeric(exclude_ALL_amount(cow)) Then exclude_ALL_amount(cow) = FormatNumber(exclude_ALL_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "exclude_ALL_hours" Then Call read_txt_array(exclude_ALL_hours, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_ALL_hours(cow) = temp_array(cow)
        '         If IsNumeric(exclude_ALL_hours(cow)) Then exclude_ALL_hours(cow) = FormatNumber(exclude_ALL_hours(the_thing), 2,,0)
        '     next
        ' End If
        If header = "exclude_SNAP_amount" Then Call read_txt_array(exclude_SNAP_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_SNAP_amount(cow) = temp_array(cow)
        '         If IsNumeric(exclude_SNAP_amount(cow)) Then exclude_SNAP_amount(cow) = FormatNumber(exclude_SNAP_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "exclude_SNAP_hours" Then Call read_txt_array(exclude_SNAP_hours, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_SNAP_hours(cow) = temp_array(cow)
        '         If IsNumeric(exclude_SNAP_hours(cow)) Then exclude_SNAP_hours(cow) = FormatNumber(exclude_SNAP_hours(the_thing), 2,,0)
        '     next
        ' End If
        If header = "exclude_CASH_amount" Then Call read_txt_array(exclude_CASH_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_CASH_amount(cow) = temp_array(cow)
        '         If IsNumeric(exclude_CASH_amount(cow)) Then exclude_CASH_amount(cow) = FormatNumber(exclude_CASH_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "exclude_CASH_hours" Then Call read_txt_array(exclude_CASH_hours, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         exclude_CASH_hours(cow) = temp_array(cow)
        '         If IsNumeric(exclude_CASH_hours(cow)) Then exclude_CASH_hours(cow) = FormatNumber(exclude_CASH_hours(the_thing), 2,,0)
        '     next
        ' End If
        If header = "SNAP_info_string" Then Call read_txt_array(SNAP_info_string, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         SNAP_info_string(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "CASH_info_string" Then Call read_txt_array(CASH_info_string, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         CASH_info_string(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "check_order" Then Call read_txt_array(check_order, info, "Number", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         check_order(cow) = temp_array(cow)
        '         If IsNumeric(check_order(cow)) Then check_order(cow) = check_order(cow)*1
        '     next
        ' End If
        If header = "view_pay_date" Then Call read_txt_array(view_pay_date, info, "Date", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         view_pay_date(cow) = temp_array(cow)
        '         If IsDate(view_pay_date(cow)) Then view_pay_date(cow) = DateAdd("d", 0, view_pay_date(cow))
        '     next
        ' End If
        If header = "frequency_issue" Then Call read_txt_array(frequency_issue, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         frequency_issue(cow) = temp_array(cow)
        '         If UCase(frequency_issue(cow)) = "TRUE" Then frequency_issue(cow) = True
        '         If UCase(frequency_issue(cow)) = "FALSE" Then frequency_issue(cow) = False
        '     next
        ' End If
        If header = "future_check" Then Call read_txt_array(future_check, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         future_check(cow) = temp_array(cow)
        '         If UCase(future_check(cow)) = "TRUE" Then future_check(cow) = True
        '         If UCase(future_check(cow)) = "FALSE" Then future_check(cow) = False
        '     next
        ' End If
        If header = "duplicate_pay_date" Then Call read_txt_array(duplicate_pay_date, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         duplicate_pay_date(cow) = temp_array(cow)
        '         If UCase(duplicate_pay_date(cow)) = "TRUE" Then duplicate_pay_date(cow) = True
        '         If UCase(duplicate_pay_date(cow)) = "FALSE" Then duplicate_pay_date(cow) = False
        '     next
        ' End If
        If header = "reason_SNAP_amt_excluded" Then Call read_txt_array(reason_SNAP_amt_excluded, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         reason_SNAP_amt_excluded(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "reason_CASH_amt_excluded" Then Call read_txt_array(reason_CASH_amt_excluded, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         reason_CASH_amt_excluded(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "pay_detail_btn" Then 'Call read_txt_array(pay_date, info, "", "^&**&^", False)
            temp_array = split(info, "^&**&^")
            for cow = 0 to UBound(temp_array)
                pay_detail_btn(cow) = 2000+cow
            next
        End If
        If header = "check_info_entered" Then Call read_txt_array(check_info_entered, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         check_info_entered(cow) = temp_array(cow)
        '         If UCase(check_info_entered(cow)) = "TRUE" Then check_info_entered(cow) = True
        '         If UCase(check_info_entered(cow)) = "FALSE" Then check_info_entered(cow) = False
        '     next
        ' End If
        If header = "bonus_check" Then Call read_txt_array(bonus_check, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         bonus_check(cow) = temp_array(cow)
        '         If UCase(bonus_check(cow)) = "TRUE" Then bonus_check(cow) = True
        '         If UCase(bonus_check(cow)) = "FALSE" Then bonus_check(cow) = False
        '     next
        ' End If
        If header = "pay_split_regular_amount" Then Call read_txt_array(pay_split_regular_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_regular_amount(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_regular_amount(cow)) Then pay_split_regular_amount(cow) = FormatNumber(pay_split_regular_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_bonus_amount" Then Call read_txt_array(pay_split_bonus_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_bonus_amount(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_bonus_amount(cow)) Then pay_split_bonus_amount(cow) = FormatNumber(pay_split_bonus_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_ot_amount" Then Call read_txt_array(pay_split_ot_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_ot_amount(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_ot_amount(cow)) Then pay_split_ot_amount(cow) = FormatNumber(pay_split_ot_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_ot_hours" Then Call read_txt_array(pay_split_ot_hours, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_ot_hours(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_ot_hours(cow)) Then pay_split_ot_hours(cow) = FormatNumber(pay_split_ot_hours(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_shift_diff_amount" Then Call read_txt_array(pay_split_shift_diff_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_shift_diff_amount(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_shift_diff_amount(cow)) Then pay_split_shift_diff_amount(cow) = FormatNumber(pay_split_shift_diff_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_tips_amount" Then Call read_txt_array(pay_split_tips_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_tips_amount(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_tips_amount(cow)) Then pay_split_tips_amount(cow) = FormatNumber(pay_split_tips_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_other_amount" Then Call read_txt_array(pay_split_other_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_other_amount(cow) = temp_array(cow)
        '         If IsNumeric(pay_split_other_amount(cow)) Then pay_split_other_amount(cow) = FormatNumber(pay_split_other_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "pay_split_other_detail" Then Call read_txt_array(pay_split_other_detail, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_split_other_detail(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "pay_excld_bonus" Then Call read_txt_array(pay_excld_bonus, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_excld_bonus(cow) = temp_array(cow)
        '         If pay_excld_bonus(cow) = "" Then pay_excld_bonus(cow) = unchecked
        '         pay_excld_bonus(cow) = pay_excld_bonus(cow) * 1
        '     next
        ' End If
        If header = "pay_excld_ot" Then Call read_txt_array(pay_excld_ot, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_excld_ot(cow) = temp_array(cow)
        '         If pay_excld_ot(cow) = "" Then pay_excld_ot(cow) = unchecked
        '         pay_excld_ot(cow) = pay_excld_ot(cow) * 1
        '     next
        ' End If
        If header = "pay_excld_shift_diff" Then Call read_txt_array(pay_excld_shift_diff, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_excld_shift_diff(cow) = temp_array(cow)
        '         If pay_excld_shift_diff(cow) = "" Then pay_excld_shift_diff(cow) = unchecked
        '         pay_excld_shift_diff(cow) = pay_excld_shift_diff(cow) * 1
        '     next
        ' End If
        If header = "pay_excld_tips" Then Call read_txt_array(pay_excld_tips, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_excld_tips(cow) = temp_array(cow)
        '         If pay_excld_tips(cow) = "" Then pay_excld_tips(cow) = unchecked
        '         pay_excld_tips(cow) = pay_excld_tips(cow) * 1
        '     next
        ' End If
        If header = "pay_excld_other" Then Call read_txt_array(pay_excld_other, info, "Checkbox", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_excld_other(cow) = temp_array(cow)
        '         If pay_excld_other(cow) = "" Then pay_excld_other(cow) = unchecked
        '         pay_excld_other(cow) = pay_excld_other(cow) * 1
        '     next
        ' End If
        If header = "split_check_string" Then Call read_txt_array(split_check_string, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         split_check_string(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "split_check_excld_string" Then Call read_txt_array(split_check_excld_string, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         split_check_excld_string(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "split_exclude_amount" Then Call read_txt_array(split_exclude_amount, info, "Amount", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         split_exclude_amount(cow) = temp_array(cow)
        '         If IsNumeric(split_exclude_amount(cow)) Then split_exclude_amount(cow) = FormatNumber(split_exclude_amount(the_thing), 2,,0)
        '     next
        ' End If
        If header = "duplct_pay_date" Then Call read_txt_array(duplct_pay_date, info, "Date", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         duplct_pay_date(cow) = temp_array(cow)
        '         If IsDate(duplct_pay_date(cow)) Then duplct_pay_date(cow) = DateAdd("d", 0, duplct_pay_date(cow))
        '     next
        ' End If
        If header = "calculated_by_ytd" Then Call read_txt_array(calculated_by_ytd, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         calculated_by_ytd(cow) = temp_array(cow)
        '         If UCase(calculated_by_ytd(cow)) = "TRUE" Then calculated_by_ytd(cow) = True
        '         If UCase(calculated_by_ytd(cow)) = "FALSE" Then calculated_by_ytd(cow) = False
        '     next
        ' End If
        If header = "ytd_calc_notes" Then Call read_txt_array(ytd_calc_notes, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         ytd_calc_notes(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "pay_detail_exists" Then Call read_txt_array(pay_detail_exists, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         pay_detail_exists(cow) = temp_array(cow)
        '         If UCase(pay_detail_exists(cow)) = "TRUE" Then pay_detail_exists(cow) = True
        '         If UCase(pay_detail_exists(cow)) = "FALSE" Then pay_detail_exists(cow) = False
        '     next
        ' End If
        If header = "combined_into_one" Then Call read_txt_array(combined_into_one, info, "Boolean", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         combined_into_one(cow) = temp_array(cow)
        '         If UCase(combined_into_one(cow)) = "TRUE" Then combined_into_one(cow) = True
        '         If UCase(combined_into_one(cow)) = "FALSE" Then combined_into_one(cow) = False
        '     next
        ' End If
        If header = "CASH_dialog_display" Then Call read_txt_array(CASH_dialog_display, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         SNAP_dialog_display(cow) = temp_array(cow)
        '     next
        ' End If
        If header = "CASH_dialog_display" Then Call read_txt_array(CASH_dialog_display, info, "String", "^&**&^", False)
        '     temp_array = split(info, "^&**&^")
        '     for cow = 0 to UBound(temp_array)
        '         CASH_dialog_display(cow) = temp_array(cow)
        '     next
        ' End If


    end sub

    public sub save_info(objTextStream)
        'record details of the class to a txt file for save_your_work

	    objTextStream.WriteLine "SET_NEW~%~%~%~%~ "
	    objTextStream.WriteLine "new_panel~%~%~%~%~" & new_panel
	    objTextStream.WriteLine "member~%~%~%~%~" & member
	    objTextStream.WriteLine "instance~%~%~%~%~" & instance
	    objTextStream.WriteLine "employer~%~%~%~%~" & employer
	    objTextStream.WriteLine "employer_with_underscores~%~%~%~%~" & employer_with_underscores
	    objTextStream.WriteLine "job_type~%~%~%~%~" & job_type
	    objTextStream.WriteLine "verif_type~%~%~%~%~" & verif_type
	    objTextStream.WriteLine "hourly_wage~%~%~%~%~" & hourly_wage
	    objTextStream.WriteLine "snap_hourly_wage~%~%~%~%~" & snap_hourly_wage
	    objTextStream.WriteLine "cash_hourly_wage~%~%~%~%~" & cash_hourly_wage
	    objTextStream.WriteLine "income_start_dt~%~%~%~%~" & income_start_dt
	    objTextStream.WriteLine "income_end_dt~%~%~%~%~" & income_end_dt
	    objTextStream.WriteLine "pay_freq~%~%~%~%~" & pay_freq
	    objTextStream.WriteLine "updated_date~%~%~%~%~" & updated_date
	    objTextStream.WriteLine "old_verif~%~%~%~%~" & old_verif
	    objTextStream.WriteLine "first_retro_check~%~%~%~%~" & first_retro_check
	    objTextStream.WriteLine "initial_month_mo~%~%~%~%~" & initial_month_mo
	    objTextStream.WriteLine "initial_month_yr~%~%~%~%~" & initial_month_yr
	    objTextStream.WriteLine "hrs_per_wk~%~%~%~%~" & hrs_per_wk
	    objTextStream.WriteLine "pay_per_hr~%~%~%~%~" & pay_per_hr
	    objTextStream.WriteLine "display_hrs_per_wk~%~%~%~%~" & display_hrs_per_wk
	    objTextStream.WriteLine "display_pay_per_hr~%~%~%~%~" & display_pay_per_hr
	    objTextStream.WriteLine "known_pay_date~%~%~%~%~" & known_pay_date
	    objTextStream.WriteLine "word_for_freq~%~%~%~%~" & word_for_freq
	    objTextStream.WriteLine "apply_to_SNAP~%~%~%~%~" & apply_to_SNAP
	    objTextStream.WriteLine "apply_to_CASH~%~%~%~%~" & apply_to_CASH
	    objTextStream.WriteLine "apply_to_HC~%~%~%~%~" & apply_to_HC
	    objTextStream.WriteLine "apply_to_GRH~%~%~%~%~" & apply_to_GRH
        objTextStream.WriteLine "prog_list~%~%~%~%~" & prog_list
	    objTextStream.WriteLine "verif_date~%~%~%~%~" & verif_date
	    objTextStream.WriteLine "verif_explain~%~%~%~%~" & verif_explain
	    objTextStream.WriteLine "selection_rsn~%~%~%~%~" & selection_rsn
	    objTextStream.WriteLine "income_excluded_cash_reason~%~%~%~%~" & income_excluded_cash_reason
	    objTextStream.WriteLine "hc_budg_notes~%~%~%~%~" & hc_budg_notes
	    objTextStream.WriteLine "hc_retro~%~%~%~%~" & hc_retro
	    objTextStream.WriteLine "EI_panel_vbYes~%~%~%~%~" & EI_panel_vbYes
	    objTextStream.WriteLine "spoke_with~%~%~%~%~" & spoke_with
	    objTextStream.WriteLine "convo_detail~%~%~%~%~" & convo_detail
	    objTextStream.WriteLine "paycheck_list_title~%~%~%~%~" & paycheck_list_title
	    objTextStream.WriteLine "excl_cash_rsn~%~%~%~%~" & excl_cash_rsn
        objTextStream.WriteLine "budget_confirmed~%~%~%~%~" & budget_confirmed
	    objTextStream.WriteLine "SNAP_list_of_excluded_pay_dates~%~%~%~%~" & SNAP_list_of_excluded_pay_dates
	    objTextStream.WriteLine "CASH_list_of_excluded_pay_dates~%~%~%~%~" & CASH_list_of_excluded_pay_dates
	    objTextStream.WriteLine "pay_weekday~%~%~%~%~" & pay_weekday
	    objTextStream.WriteLine "income_received~%~%~%~%~" & income_received
	    objTextStream.WriteLine "all_pay_in_app_month~%~%~%~%~" & all_pay_in_app_month
	    objTextStream.WriteLine "there_are_counted_checks~%~%~%~%~" & there_are_counted_checks
	    objTextStream.WriteLine "actual_checks_provided~%~%~%~%~" & actual_checks_provided
	    objTextStream.WriteLine "anticipated_income_provided~%~%~%~%~" & anticipated_income_provided
	    objTextStream.WriteLine "missing_checks_list~%~%~%~%~" & missing_checks_list
	    objTextStream.WriteLine "issues_with_frequency~%~%~%~%~" & issues_with_frequency
	    objTextStream.WriteLine "default_start_date~%~%~%~%~" & default_start_date
	    objTextStream.WriteLine "bimonthly_first~%~%~%~%~" & bimonthly_first
	    objTextStream.WriteLine "bimonthly_second~%~%~%~%~" & bimonthly_second
	    objTextStream.WriteLine "ignore_antic~%~%~%~%~" & ignore_antic
	    objTextStream.WriteLine "pick_one~%~%~%~%~" & pick_one
	    objTextStream.WriteLine "days_to_add~%~%~%~%~" & days_to_add
	    objTextStream.WriteLine "months_to_add~%~%~%~%~" & months_to_add
	    objTextStream.WriteLine "first_check~%~%~%~%~" & first_check
	    objTextStream.WriteLine "last_check~%~%~%~%~" & last_check
	    objTextStream.WriteLine "order_ubound~%~%~%~%~" & order_ubound
	    objTextStream.WriteLine "checks_exist~%~%~%~%~" & checks_exist
	    objTextStream.WriteLine "estimate_exists~%~%~%~%~" & estimate_exists
        objTextStream.WriteLine "gross_max_string_len~%~%~%~%~" & gross_max_string_len
	    objTextStream.WriteLine "total_check_count~%~%~%~%~" & total_check_count
	    objTextStream.WriteLine "total_gross_amount~%~%~%~%~" & total_gross_amount
	    objTextStream.WriteLine "total_hours~%~%~%~%~" & total_hours
	    objTextStream.WriteLine "ave_hrs_per_pay~%~%~%~%~" & ave_hrs_per_pay
	    objTextStream.WriteLine "ave_inc_per_pay~%~%~%~%~" & ave_inc_per_pay
	    objTextStream.WriteLine "monthly_income~%~%~%~%~" & monthly_income
	    objTextStream.WriteLine "snap_check_count~%~%~%~%~" & snap_check_count
	    objTextStream.WriteLine "snap_budgeted_total~%~%~%~%~" & snap_budgeted_total
	    objTextStream.WriteLine "snap_hours~%~%~%~%~" & snap_hours
	    objTextStream.WriteLine "snap_ave_hrs_per_pay~%~%~%~%~" & snap_ave_hrs_per_pay
	    objTextStream.WriteLine "snap_ave_inc_per_pay~%~%~%~%~" & snap_ave_inc_per_pay
	    objTextStream.WriteLine "snap_hrs_per_wk~%~%~%~%~" & snap_hrs_per_wk
	    objTextStream.WriteLine "SNAP_monthly_income~%~%~%~%~" & SNAP_monthly_income
	    objTextStream.WriteLine "cash_check_count~%~%~%~%~" & cash_check_count
	    objTextStream.WriteLine "cash_budgeted_total~%~%~%~%~" & cash_budgeted_total
	    objTextStream.WriteLine "cash_hours~%~%~%~%~" & cash_hours
	    objTextStream.WriteLine "cash_ave_hrs_per_pay~%~%~%~%~" & cash_ave_hrs_per_pay
	    objTextStream.WriteLine "cash_ave_inc_per_pay~%~%~%~%~" & cash_ave_inc_per_pay
	    objTextStream.WriteLine "cash_hrs_per_wk~%~%~%~%~" & cash_hrs_per_wk
	    objTextStream.WriteLine "CASH_monthly_income~%~%~%~%~" & CASH_monthly_income
	    objTextStream.WriteLine "numb_months~%~%~%~%~" & numb_months
	    objTextStream.WriteLine "reg_non_monthly~%~%~%~%~" & reg_non_monthly
	    objTextStream.WriteLine "update_this_month~%~%~%~%~" & update_this_month
	    objTextStream.WriteLine "months_updated~%~%~%~%~" & months_updated
	    objTextStream.WriteLine "income_lumped_mo~%~%~%~%~" & income_lumped_mo
	    objTextStream.WriteLine "lump_reason~%~%~%~%~" & lump_reason
	    objTextStream.WriteLine "act_checks_lumped~%~%~%~%~" & act_checks_lumped
	    objTextStream.WriteLine "est_checks_lumped~%~%~%~%~" & est_checks_lumped
	    objTextStream.WriteLine "lump_gross~%~%~%~%~" & lump_gross
	    objTextStream.WriteLine "lump_hrs~%~%~%~%~" & lump_hrs
	    objTextStream.WriteLine "mo_w_more_5_chcks~%~%~%~%~" & mo_w_more_5_chcks
	    objTextStream.WriteLine "update_future~%~%~%~%~" & update_future
        objTextStream.WriteLine "updates_to_display~%~%~%~%~" & updates_to_display
	    objTextStream.WriteLine "cash_array_info_exists~%~%~%~%~" & cash_array_info_exists

	    If IsArray(cash_anticipated_pay_array) Then objTextStream.WriteLine "cash_anticipated_pay_array~%~%~%~%~" & join(cash_anticipated_pay_array, "|")
	    If IsArray(snap_anticipated_pay_array) Then objTextStream.WriteLine "snap_anticipated_pay_array~%~%~%~%~" & join(snap_anticipated_pay_array, "|")
        objTextStream.WriteLine "expected_check_array~%~%~%~%~" & join(expected_check_array, "|")
        objTextStream.WriteLine "this_month_checks_array~%~%~%~%~" & join(this_month_checks_array, "|")
        objTextStream.WriteLine "retro_month_checks_array~%~%~%~%~" & join(retro_month_checks_array, "|")


        objTextStream.WriteLine "CASH_ARRAY~%~%~%~%~" & UBound(cash_info_cash_mo_yr)

        objTextStream.WriteLine "cash_info_cash_mo_yr~%~%~%~%~" & join(cash_info_cash_mo_yr, "^&**&^")
        objTextStream.WriteLine "cash_info_retro_mo_yr~%~%~%~%~" & join(cash_info_retro_mo_yr, "^&**&^")
        objTextStream.WriteLine "cash_info_retro_updtd~%~%~%~%~" & join(cash_info_retro_updtd, "^&**&^")
        objTextStream.WriteLine "cash_info_prosp_updtd~%~%~%~%~" & join(cash_info_prosp_updtd, "^&**&^")
        objTextStream.WriteLine "cash_info_mo_retro_pay~%~%~%~%~" & join(cash_info_mo_retro_pay, "^&**&^")
        objTextStream.WriteLine "cash_info_mo_retro_hrs~%~%~%~%~" & join(cash_info_mo_retro_hrs, "^&**&^")
        objTextStream.WriteLine "cash_info_mo_prosp_pay~%~%~%~%~" & join(cash_info_mo_prosp_pay, "^&**&^")
        objTextStream.WriteLine "cash_info_mo_prosp_hrs~%~%~%~%~" & join(cash_info_mo_prosp_hrs, "^&**&^")

        objTextStream.WriteLine "PAYCHECK_ARRAY~%~%~%~%~" & UBound(pay_date)

        objTextStream.WriteLine "pay_date~%~%~%~%~" & join(pay_date, "^&**&^")
        objTextStream.WriteLine "gross_amount~%~%~%~%~" & join(gross_amount, "^&**&^")
        objTextStream.WriteLine "hours~%~%~%~%~" & join(hours, "^&**&^")
        objTextStream.WriteLine "exclude_entirely~%~%~%~%~" & join(exclude_entirely, "^&**&^")
        objTextStream.WriteLine "exclude_from_SNAP~%~%~%~%~" & join(exclude_from_SNAP, "^&**&^")
        objTextStream.WriteLine "exclude_from_CASH~%~%~%~%~" & join(exclude_from_CASH, "^&**&^")
        objTextStream.WriteLine "reason_to_exclude~%~%~%~%~" & join(reason_to_exclude, "^&**&^")
        objTextStream.WriteLine "exclude_ALL_amount~%~%~%~%~" & join(exclude_ALL_amount, "^&**&^")
        objTextStream.WriteLine "exclude_ALL_hours~%~%~%~%~" & join(exclude_ALL_hours, "^&**&^")
        objTextStream.WriteLine "exclude_SNAP_amount~%~%~%~%~" & join(exclude_SNAP_amount, "^&**&^")
        objTextStream.WriteLine "exclude_SNAP_hours~%~%~%~%~" & join(exclude_SNAP_hours, "^&**&^")
        objTextStream.WriteLine "exclude_CASH_amount~%~%~%~%~" & join(exclude_CASH_amount, "^&**&^")
        objTextStream.WriteLine "exclude_CASH_hours~%~%~%~%~" & join(exclude_CASH_hours, "^&**&^")
        objTextStream.WriteLine "SNAP_info_string~%~%~%~%~" & join(SNAP_info_string, "^&**&^")
        objTextStream.WriteLine "CASH_info_string~%~%~%~%~" & join(CASH_info_string, "^&**&^")
        objTextStream.WriteLine "check_order~%~%~%~%~" & join(check_order, "^&**&^")
        objTextStream.WriteLine "view_pay_date~%~%~%~%~" & join(view_pay_date, "^&**&^")
        objTextStream.WriteLine "frequency_issue~%~%~%~%~" & join(frequency_issue, "^&**&^")
        objTextStream.WriteLine "future_check~%~%~%~%~" & join(future_check, "^&**&^")
        objTextStream.WriteLine "duplicate_pay_date~%~%~%~%~" & join(duplicate_pay_date, "^&**&^")
        objTextStream.WriteLine "reason_SNAP_amt_excluded~%~%~%~%~" & join(reason_SNAP_amt_excluded, "^&**&^")
        objTextStream.WriteLine "reason_CASH_amt_excluded~%~%~%~%~" & join(reason_CASH_amt_excluded, "^&**&^")
        objTextStream.WriteLine "pay_detail_btn~%~%~%~%~" & join(pay_detail_btn, "^&**&^")
        objTextStream.WriteLine "check_info_entered~%~%~%~%~" & join(check_info_entered, "^&**&^")
        objTextStream.WriteLine "bonus_check~%~%~%~%~" & join(bonus_check, "^&**&^")
        objTextStream.WriteLine "pay_split_regular_amount~%~%~%~%~" & join(pay_split_regular_amount, "^&**&^")
        objTextStream.WriteLine "pay_split_bonus_amount~%~%~%~%~" & join(pay_split_bonus_amount, "^&**&^")
        objTextStream.WriteLine "pay_split_ot_amount~%~%~%~%~" & join(pay_split_ot_amount, "^&**&^")
        objTextStream.WriteLine "pay_split_ot_hours~%~%~%~%~" & join(pay_split_ot_hours, "^&**&^")
        objTextStream.WriteLine "pay_split_shift_diff_amount~%~%~%~%~" & join(pay_split_shift_diff_amount, "^&**&^")
        objTextStream.WriteLine "pay_split_tips_amount~%~%~%~%~" & join(pay_split_tips_amount, "^&**&^")
        objTextStream.WriteLine "pay_split_other_amount~%~%~%~%~" & join(pay_split_other_amount, "^&**&^")
        objTextStream.WriteLine "pay_split_other_detail~%~%~%~%~" & join(pay_split_other_detail, "^&**&^")
        objTextStream.WriteLine "pay_excld_bonus~%~%~%~%~" & join(pay_excld_bonus, "^&**&^")
        objTextStream.WriteLine "pay_excld_ot~%~%~%~%~" & join(pay_excld_ot, "^&**&^")
        objTextStream.WriteLine "pay_excld_shift_diff~%~%~%~%~" & join(pay_excld_shift_diff, "^&**&^")
        objTextStream.WriteLine "pay_excld_tips~%~%~%~%~" & join(pay_excld_tips, "^&**&^")
        objTextStream.WriteLine "pay_excld_other~%~%~%~%~" & join(pay_excld_other, "^&**&^")
        objTextStream.WriteLine "split_check_string~%~%~%~%~" & join(split_check_string, "^&**&^")
        objTextStream.WriteLine "split_check_excld_string~%~%~%~%~" & join(split_check_excld_string, "^&**&^")
        objTextStream.WriteLine "split_exclude_amount~%~%~%~%~" & join(split_exclude_amount, "^&**&^")
        objTextStream.WriteLine "duplct_pay_date~%~%~%~%~" & join(duplct_pay_date, "^&**&^")
        objTextStream.WriteLine "calculated_by_ytd~%~%~%~%~" & join(calculated_by_ytd, "^&**&^")
        objTextStream.WriteLine "ytd_calc_notes~%~%~%~%~" & join(ytd_calc_notes, "^&**&^")
        objTextStream.WriteLine "pay_detail_exists~%~%~%~%~" & join(pay_detail_exists, "^&**&^")
        objTextStream.WriteLine "combined_into_one~%~%~%~%~" & join(combined_into_one, "^&**&^")
        objTextStream.WriteLine "SNAP_dialog_display~%~%~%~%~" & join(SNAP_dialog_display, "^&**&^")
        objTextStream.WriteLine "CASH_dialog_display~%~%~%~%~" & join(CASH_dialog_display, "^&**&^")
	    objTextStream.WriteLine "NEXT_CLASS~%~%~%~%~ "

    end sub

	public sub select_budget_option()
        'If both actual and expected income information is provided, user must select which to use for budgeting

		If there_are_counted_checks = FALSE AND actual_checks_provided = TRUE Then
			pick_one = use_estimate
		ElseIf actual_checks_provided = TRUE AND anticipated_income_provided = TRUE Then
			'CHOOSE CORRECT METHOD Dialog - select which (actual or anticipated) income information to budget and explain
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 196, 165, "Reasonably Expected to Continue"
				OptionGroup RadioGroup1
					RadioButton 25, 70, 130, 10, "Use the actual check amounts/dates", use_actual_income
					RadioButton 25, 85, 130, 10, "Use the anticipated hours/wage", use_anticipated_income
				EditBox 10, 125, 180, 15, selection_rsn
				ButtonGroup ButtonPressed
					OkButton 140, 145, 50, 15
				Text 10, 10, 185, 35, "Both Actual Income and Anticipated Income have been listed for a SNAP case. Since both have been reported, both will be case noted. For entering information to the PIC, one option should be selected."
				GroupBox 5, 55, 185, 45, "Which is the best estimation of anticipated income?"
				Text 10, 110, 185, 10, "Explain why this is the best estimation of future income:"
			EndDialog

			Do
				sm_err_msg = ""

				Dialog Dialog1      'one of the easiest dialogs in this script
				save_your_work

				selection_rsn = trim(selection_rsn)
				If use_actual_income = checked Then selection_pick = "ACTUAL LIST OF CHECKS."
				If use_anticipated_income = checked Then selection_pick = "INCOME ESTIMATED FROM HOURS AND RATE OF PAY."

				If selection_rsn = "" Then sm_err_msg = sm_err_msg & vbNewLine & "* Enter explanation of why the best way to determine future income is to use " & selection_pick
				If len(selection_rsn) < 10 Then sm_err_msg = sm_err_msg & vbNewLine & "* Explanation is not sufficient to adequately case note information about budget. Expand."

				If sm_err_msg <> "" Then MsgBox "** Please Resolve before Continuting **" & vbNewLine & sm_err_msg
			Loop until sm_err_msg = ""

			'Setting selections based on the choice made
			If use_actual_income = checked Then
				pick_one = use_actual
				ignore_antic = TRUE
			End If
			If use_anticipated_income = checked Then
				pick_one = use_estimate
			End If

			'https://www.dhssir.cty.dhs.state.mn.us/MAXIS/trntl/snap/SNAP_Anticipating_Income.pdf - this is all about why we have to pick
		Else        'if it isn't both then we don't need the dialog and the script sets the selection based on entry information
			If actual_checks_provided = TRUE Then pick_one = use_actual
			If anticipated_income_provided = TRUE Then pick_one = use_estimate
		End If
        save_your_work
	end sub

    public sub set_programs()
        'just a list of programs based on checkboxes
        prog_list = ""
        If apply_to_SNAP = checked Then prog_list = prog_list & "/SNAP"       'setting the header programs
        If apply_to_CASH = checked Then prog_list = prog_list & "/CASH"
        If apply_to_HC = checked Then prog_list = prog_list & "/HC"
        If apply_to_GRH = checked Then prog_list = prog_list & "/GRH"

        If left(prog_list, 1) = "/" Then prog_list = right(prog_list, len(prog_list)-1)
    end sub

    public sub update_cash_pic()
        'enter information on CASH PIC

        If apply_to_GRH = checked or apply_to_CASH = checked  Then
            STATS_manualtime = STATS_manualtime + 145
            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "*** GRH Budget Update ***" & vbNewLine & "---- PIC ----"
            EMWriteScreen "X", 19, 71               'opening the GRH PIC
            transmit


            EMWriteScreen "      ", 6, 70
            EMWriteScreen "        ", 7, 70
            EMWriteScreen "        ", 11, 70

            list_row = 7                    'here we clear the PIC of all previous data
            beg_of_list_check = ""
            Do
                EMWriteScreen "  ", list_row, 9
                EMWriteScreen "  ", list_row, 12
                EMWriteScreen "  ", list_row, 15
                EMWriteScreen "        ", list_row, 21
                EMWriteScreen "      ", list_row, 32
                list_row = list_row + 1

                EMReadScreen next_list_item, 29, list_row, 9
            Loop until list_row = 17 or next_list_item = "__ __ __    ________   ______"

            Call create_mainframe_friendly_date(date, 3, 30, "YY")                     'enter the current date in date of calculation field
            EMWriteScreen left(pay_freq, 1), 3, 63        'enter the pay frequency code only in the correct field
            updates_to_display = updates_to_display & vbNewLine & "Date of Calculation: " & date & "  Pay Frequency: " & pay_freq

            If pick_one = use_estimate Then           'if use estimate was shosen, just need the hourly wage and pay per hour
                updates_to_display = updates_to_display & vbNewLine & "Estimate: Hourly Wage - $" & hourly_wage & "/hr - " & hrs_per_wk & " hrs/wk" & vbNewLine
                EMWriteScreen hrs_per_wk, 6, 70
                EMWriteScreen hourly_wage, 7, 70
                If IsNumeric(hrs_per_wk) = True Then
                    pic_grh_ave_hrs_READ = hrs_per_wk * 1
                Else
                    pic_grh_ave_hrs_READ = 0
                End If
            End If
            If pick_one = use_actual Then             'if we are using actual, we need to put some real amounts in here
                updates_to_display = updates_to_display & vbNewLine & "Actual Pay: Date - "
                running_hours = 0

                list_row = 7                'entering all checks on the PIC that were provided - no exclusions or reductions
                For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
                    For all_income = 0 to UBound(pay_date)                  'then loop through all of the income information
                        If check_order(all_income) = order_number Then
                            If NOT exclude_entirely(all_income) and exclude_from_CASH(all_income) = unchecked  Then
                                Call create_mainframe_friendly_date(view_pay_date(all_income), list_row, 9, "YY")
                                gross_amount(all_income) = FormatNumber(gross_amount(all_income), 2, -1, 0, 0)
                                calc_exclude = 0
                                If IsNumeric(exclude_ALL_amount(all_income)) Then calc_exclude = calc_exclude + exclude_ALL_amount(all_income)
                                If IsNumeric(split_exclude_amount(all_income)) Then calc_exclude = calc_exclude + split_exclude_amount(all_income)
                                If IsNumeric(exclude_CASH_amount(all_income)) Then calc_exclude = calc_exclude + exclude_CASH_amount(all_income)
                                net_amount = gross_amount(all_income) - calc_exclude      'taking out excluded amounts
                                net_amount = FormatNumber(net_amount, 2, -1, 0, 0)
                                EMWriteScreen net_amount, list_row, 21
                                running_hours = running_hours + hours(all_income)
                                updates_to_display = updates_to_display & view_pay_date(all_income) & " - $" & net_amount & " - " & hours(all_income) & " hrs." & vbNewLine

                                list_row = list_row + 1
                            End If
                        End If
                    next
                next
                numb_of_chcks = list_row - 7
                If list_row < 17 Then
                    For clear_row = list_row to 16
                        EMWriteScreen "  ", clear_row, 9
                        EMWriteScreen "  ", clear_row, 12
                        EMWriteScreen "  ", clear_row, 15
                        EMWriteScreen "        ", clear_row, 21
                        EMWriteScreen "      ", clear_row, 32
                    Next
                End If
                pic_grh_ave_hrs_READ = running_hours/numb_of_chcks
            End If

            read_calculation = True
            Do
                transmit            'saving the PIC
                If read_calculation = True Then
                    read_calculation = False

                    EMReadScreen pic_grh_ave_hrs_READ,      10, 15, 68
                    EMReadScreen pic_grh_ave_paycheck_READ, 10, 16, 68
                    EMReadScreen pic_grh_prosp_mo_READ,     10, 17, 68

                    pic_grh_ave_hrs_READ = trim(pic_grh_ave_hrs_READ)
                    pic_grh_ave_paycheck_READ = trim(pic_grh_ave_paycheck_READ)
                    pic_grh_prosp_mo_READ = trim(pic_grh_prosp_mo_READ)
                End If
                EMReadScreen escape_route, 26, 18, 5
                If escape_route = "CHANGE DATA OR PF3 TO EXIT" Then PF3
                EMReadScreen pic_menu, 35, 2, 25
            Loop until pic_menu <> "CASH Prospective Income Calculation"
        End If
    end sub

    public sub update_hc_estimate()
        'Enter information on HC Income Estimate PIC
        If apply_to_HC = checked Then         'now on to the health care
            STATS_manualtime = STATS_manualtime + 140

            If MAXIS_footer_month = CM_plus_1_mo AND MAXIS_footer_year = CM_plus_1_yr Then          'if we are in current month plus one, we need to update the HC Income Estimate Pop-up
                EMWriteScreen "X", 19, 48           'opening the HC Inc Est
                transmit

                EMWriteScreen "        ", 11, 63        'blanking it out
                ave_inc_per_pay = FormatNumber(ave_inc_per_pay, 2, -1, 0, 0)
                EMWriteScreen ave_inc_per_pay, 11, 63         'writing it in

                Do
                    transmit            'saving the PIC
                    EMReadScreen escape_route, 26, 16, 24
                    If escape_route = "CHANGE DATA OR PF3 TO EXIT" Then PF3
                    EMReadScreen pic_menu, 18, 9, 43
                Loop until pic_menu <> "HC Income Estimate"
            End If
        End If
    end sub

	public sub update_job_detail()
        'This is the dialog with the list of checks and other details about the job

		Do
			all_info_entered = True
			Do
				call evaluate_checks()

				err_msg = ""
				dlg_factor = UBound(pay_date)
				dlg_hgt_fact = dlg_factor
				If dlg_hgt_fact = 0 Then dlg_hgt_fact = 1

				'ENTER PAY Dialog - dynamic dialog to enter job checks or anticipated amounts
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 765, (dlg_hgt_fact * 10) + 155, "Enter ALL Paychecks Received"
					ButtonGroup ButtonPressed
					Text 10, 10, 265, 10, "JOBS " & member & " " & instance & " - " & employer
					Text 200, 15, 40, 10, "Start Date:"
					EditBox 235, 10, 50, 15, income_start_dt
					Text 295, 15, 50, 10, "Income Type:"
					DropListBox 345, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program"+chr(9)+"N - Census Income", job_type
					GroupBox 455, 5, 145, 25, "Apply Income to Programs:"
					CheckBox 465, 15, 30, 10, "SNAP", apply_to_SNAP
					CheckBox 500, 15, 30, 10, "CASH", apply_to_CASH
					CheckBox 535, 15, 20, 10, "HC", apply_to_HC
					CheckBox 560, 15, 30, 10, "GRH", apply_to_GRH
					Text 615, 15, 90, 10, "Date verification received:"
					EditBox 710, 10, 50, 15, verif_date
					Text 5, 40, 60, 10, "JOBS Verif Code:"
					DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd"+chr(9)+"? - EXPEDITED SNAP ONLY", verif_type
					Text 175, 40, 155, 10, "additional detail of verification received:"
					EditBox 310, 35, 290, 15, verif_explain
					Text 625, 40, 50, 10, "Pay Frequency"
					DropListBox 675, 35, 85, 45, ""+chr(9)+"1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", pay_freq

					PushButton 605, 35, 15, 15, "!", pay_frequency_tips_and_tricks_btn
					' PushButton 5, 50, 15, 15, "!", listing_checks_tips_and_tricks_btn

					GroupBox 5, 55, 750, (dlg_factor * 10) + 35, "Check Details"
					If not checks_exist Then
						Text 15, 70, 150, 10, "No Actual Checks Entered"
					Else
						Text 115, 65, 60, 10, "Pay Date:"
						Text 160, 65, 50, 10, "Gross Amount:"
						Text 210, 65, 25, 10, "Hours:"
						Text 245, 65, 75, 10, "Budget Details:"

						y_pos = 0     'this is how we move things down in dynamic dialogs
						For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
							For all_checks = 0 to UBound(pay_date)
								If check_order(all_checks) = order_number Then
									PushButton 15, (y_pos * 10) + 75, 90, 10, "Update Check Details", pay_detail_btn(all_checks)
									Text 115, (y_pos * 10) + 75, 50, 10, pay_date(all_checks)
									Text 160, (y_pos * 10) + 75, 45, 10, "$ " & gross_amount(all_checks)
									Text 210, (y_pos * 10) + 75, 30, 10, hours(all_checks)
									If exclude_entirely(check_count) Then
										Text 245, (y_pos * 10) + 75, 500, 10, "Entire Check Excluded from both Cash and SNAP. Reason: " & reason_to_exclude(all_checks)
									ElseIf bonus_check(all_checks) Then
										Text 245, (y_pos * 10) + 75, 500, 10, "BONUS CHECK: Entire Check Excluded from both Cash and SNAP."
									ElseIf split_check_string(all_checks) <> "" Then
										Text 245, (y_pos * 10) + 75, 500, 10, split_check_string(all_checks)
									ElseIf SNAP_info_string(all_checks) = "Entire check counted for SNAP." and CASH_info_string(all_checks) = "Entire check counted for CASH." Then
										Text 245, (y_pos * 10) + 75, 500, 10, "Entire check counted for all programs."
									Else
										Text 245, (y_pos * 10) + 75, 250, 10, "SNAP: " & SNAP_info_string(all_checks)
										Text 500, (y_pos * 10) + 75, 250, 10, "CASH: " & CASH_info_string(all_checks)
									End If
									' Text 355, (y_pos * 10) + 75, 45, 10, pay_date(all_checks)
									' Text 410, (y_pos * 10) + 75, 185, 10, pay_date(all_checks)
									' Text 600, (y_pos * 10) + 75, 60, 10, pay_date(all_checks)
									y_pos = y_pos + 1
								End If
							Next
						Next
					End If
					If dlg_factor = 0 Then dlg_factor = 1
					PushButton 5, (dlg_factor * 10) + 100, 75, 13, "Add More Checks", add_another_check_btn
					PushButton 85, (dlg_factor * 10) + 100, 200, 13, "Insert Check using YTD calculation from surrounding checks", ytd_calculator_btn

					Text 22, (dlg_factor * 10) + 130, 185, 20, "List ALL known/reported/verified checks with gross amounts, even if not used to create a prospective budget."
					Text 545, (dlg_factor * 10) + 105, 85, 10, "Initial Month to Update:"
					EditBox 625, (dlg_factor * 10) + 100, 15, 15, initial_month_mo
					EditBox 645, (dlg_factor * 10) + 100, 15, 15, initial_month_yr
					CheckBox 670, (dlg_factor * 10) + 105, 120, 10, "Update Future Months", update_future

					GroupBox 235, (dlg_factor * 10) + 120, 275, 30, "Anticipated Income:"
					Text 240, (dlg_factor * 10) + 135, 50, 10, "Rate of Pay/Hr"
					EditBox 290, (dlg_factor * 10) + 130, 30, 15, display_pay_per_hr
					Text 325, (dlg_factor * 10) + 135, 35, 10, "Hours/Wk"
					EditBox 360, (dlg_factor * 10) + 130, 20, 15, display_hrs_per_wk
					Text 390, (dlg_factor * 10) + 135, 65, 10, "Known Pay Date"
					EditBox 450, (dlg_factor * 10) + 130, 50, 15, known_pay_date

					PushButton 5, (dlg_factor * 10) + 130, 15, 15, "!", list_all_checks_tips_and_checks_btn
					PushButton 525, (dlg_factor * 10) + 100, 15, 15, "!", initial_month_tips_and_tricks_btn
					OkButton 710, (dlg_factor * 10) + 135, 50, 15
				EndDialog

				Dialog Dialog1
				cancel_confirmation     'there is no cancel button but this will make sure that if the 'X' is pressed the worker has a way out
				save_your_work

                pay_per_hr = display_pay_per_hr
                hrs_per_wk = display_hrs_per_wk

				call evaluate_job_info
				call job_info_error_handling(err_msg)

				For all_checks = 0 to UBound(pay_detail_btn)
					If ButtonPressed = pay_detail_btn(all_checks) Then
						err_msg = "LOOP" & err_msg
						call check_details_dialog(True, all_checks, ButtonPressed)
						If ButtonPressed = save_details_btn Then
							check_info_entered(all_checks) = True
						End If
						If ButtonPressed = delete_check_btn Then
							Call delete_one_check(all_checks)
						End If
						Exit For
					End If
				Next

				If ButtonPressed = ytd_calculator_btn Then
					If actual_checks_provided = False Then
						err_msg = err_msg & vbCr & "* You selected to have the script complete a YTD calculation, but there are no check detials entered and this functionality cannot operate."
					ElseIf pay_freq <> "" Then
						Call create_expected_check_array
						Call find_missing_checks
						Call ytd_calculator(err_msg)
						Call order_checks
					End If
				End If
				If ButtonPressed = add_another_check_btn Then Call add_check
				If ButtonPressed = pay_frequency_tips_and_tricks_btn Then tips_and_tricks_msg = MsgBox(pay_freq_msg_text, vbInformation, "Tips and Tricks")
				If ButtonPressed = list_all_checks_tips_and_checks_btn Then tips_and_tricks_msg = MsgBox(all_checks_msg_text, vbInformation, "Tips and Tricks")
				If ButtonPressed = initial_month_tips_and_tricks_btn Then tips_and_tricks_msg = MsgBox(initial_month_msg_text, vbInformation, "Tips and Tricks")
				If ButtonPressed <> -1 and ButtonPressed <> ytd_calculator_btn Then err_msg = "LOOP" & err_msg

				If err_msg <> "" AND left(err_msg, 4) <> "LOOP" then MsgBox "Please resolve before continuing:" & vbNewLine & err_msg      'shoing the error message if there is one

			Loop until err_msg = ""

			If actual_checks_provided = TRUE Then
				Call order_checks
				Call create_expected_check_array
				Call find_missing_checks
			End If

			Call select_budget_option

			If actual_checks_provided = TRUE Then
				If missing_checks_list <> "" Then       'if there were any missing checks found
					If left(missing_checks_list, 1) = "~" Then missing_checks_list = right(missing_checks_list, len(missing_checks_list) - 1)       'create an array of the missing checks
					missing_checks_array = split(missing_checks_list, "~")
					loop_to_add_missing_checks = TRUE       'these are set to show the ENTER PAY Dialog again without going to he CONFIRM BUDGET Dialog
					review_small_dlg = TRUE

					'telling the worker why we are going back
					MsgBox "*** It appears there are checks missing ***" & vbNewLine & vbNewLine & "All checks need to be entered to have a correct budget. If there are pay dates between the first and last date entered that were not included, include them now. If the pay was $0, list $0 income."

					For each check_missed in missing_checks_array        'these missing dates get added to the LIST_OF_INCOME_ARRAY automatically
						pay_item = UBound(pay_date)
						If pay_date(pay_item) <> "" Then pay_item = pay_item + 1
						Call resize_check_list(pay_item)
						pay_date(pay_item) = check_missed
						pay_detail_btn(pay_item) = 2000+pay_item
						duplicate_pay_date(pay_item) = False
						check_info_entered(pay_item) = True
						future_check(pay_item) = False
						If DatePart("m", fs_appl_date) = DatePart("m", pay_date(pay_item)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", pay_date(pay_item)) Then   'if the paydate is in the application month
							If DateDiff("d", date, pay_date(pay_item)) > 0 Then future_check(pay_item) = TRUE   'this is a future check
						End If

						Call check_details_dialog(True, pay_item, ButtonPressed)
						If ButtonPressed = save_details_btn Then
							check_info_entered(pay_item) = True
						End If
						If ButtonPressed = delete_check_btn Then
							Call delete_one_check(pay_item)
						End If
					Next
					all_info_entered = False
					Call order_checks
				End If
			End If
		Loop until all_info_entered = True

		call determine_weekday_of_pay
		call evaluate_frequency
		call calculate_totals
	end sub

    public sub update_main_panel_checks()
        'enter information to the main panel

        jobs_row = 12               'first we blank out what is already there
        jobs_col = 25
        Do
            If jobs_col = 38 Then
                EMWriteScreen "        ", jobs_row, jobs_col
            ElseIf jobs_col = 67 Then
                EMWriteScreen "        ", jobs_row, jobs_col
            Else
                EMWriteScreen "  ", jobs_row, jobs_col
            End If

            jobs_col = jobs_col + 3
            If jobs_col = 34 then jobs_col = 38
            If jobs_col = 41 then jobs_col = 54
            If jobs_col = 63 then jobs_col = 67
            If jobs_col = 70 Then
                jobs_col = 25
                jobs_row = jobs_row + 1
            End If
        Loop until jobs_row = 17

        jobs_row = 12               'now we are going to enter the new information
        total_hours = 0
        income_items_used = "~"

        updates_to_display = updates_to_display & "-- Main JOBS panel --"

        checks_in_month = 0
        check_to_enter = "~"
        For all_income = 0 to UBound(pay_date)
            If pay_date(all_income) <> "" Then
                If DatePart("m", view_pay_date(all_income)) = DatePart("m", this_month_checks_array(0)) Then
                    If InStr(check_to_enter, "~"& view_pay_date(all_income) &"~") = 0 Then
                        checks_in_month = checks_in_month + 1
                        check_to_enter = check_to_enter & view_pay_date(all_income) & "~"
                    End If
                End If
            End If
        Next

        If checks_in_month > 5 or code_for_six_month_reporting_workaround = True Then
            If checks_in_month > 5 Then mo_w_more_5_chcks = mo_w_more_5_chcks & " " & MAXIS_footer_month & "/" & MAXIS_footer_year

            EMWriteScreen MAXIS_footer_month, 12, 54
            EMWriteScreen "01", 12, 57
            EMWriteScreen MAXIS_footer_year, 12, 60
            If job_start_month = MAXIS_footer_month AND job_start_year = MAXIS_footer_year Then Call write_date(income_start_dt(ei_panel), "MM DD YY", 12, 54)

            updates_to_display = updates_to_display & vbNewLine & "GROUPING ALL " & MAXIS_footer_month & "/" & MAXIS_footer_year & "checks"

            this_month_total_gross = 0
            If code_for_six_month_reporting_workaround = True Then
                If cash_workaround_ga = True Then
                    this_month_total_gross = pic_grh_prosp_mo_READ
                    'PROBLEM!!! GRH PIC DOESN'T GIVE HOURS
                    If left(pay_freq, 1) = "1" Then workaround_multiplier = 1
                    If left(pay_freq, 1) = "2" Then workaround_multiplier = 2
                    If left(pay_freq, 1) = "3" Then workaround_multiplier = 2.15
                    If left(pay_freq, 1) = "4" Then workaround_multiplier = 4.3
                    total_hours = pic_grh_ave_hrs_READ * workaround_multiplier
                    updates_to_display = updates_to_display & vbNewLine & "Six-Month Workaround Coded for GA: GRH PIC Monthly Income $ " & pic_grh_prosp_mo_READ & ", Average Hours: " & pic_grh_ave_hrs_READ
                End If
                If UH_SNAP = TRUE or cash_workaround_mfip = True Then
                    this_month_total_gross = pic_fs_prosp_mo_READ
                    If left(pay_freq, 1) = "1" Then workaround_multiplier = 1
                    If left(pay_freq, 1) = "2" Then workaround_multiplier = 2
                    If left(pay_freq, 1) = "3" Then workaround_multiplier = 2.15
                    If left(pay_freq, 1) = "4" Then workaround_multiplier = 4.3
                    If IsNumeric(pic_fs_ave_hrs_READ) = True Then
                        pic_fs_ave_hrs_READ = pic_fs_ave_hrs_READ * 1
                    Else
                        pic_fs_ave_hrs_READ = 0
                    End If
                    total_hours = pic_fs_ave_hrs_READ * workaround_multiplier
                    updates_to_display = updates_to_display & vbNewLine & "Six-Month Workaround Coded for UHFS/MFIP: SNAP PIC Monthly Income $ " & pic_fs_prosp_mo_READ & ", Average Hours: " & pic_fs_ave_hrs_READ
                End If
            Else
                For all_income = 0 to UBound(pay_date)
                    If pay_date(all_income) <> "" Then
                        If DatePart("m", view_pay_date(all_income)) = DatePart("m", this_month_checks_array(0)) Then
                            this_month_total_gross = this_month_total_gross + gross_amount(all_income)
                            total_hours = total_hours + hours(all_income)             'running total of hours
                            updates_to_display = updates_to_display & vbNewLine & "Date - " & view_pay_date(all_income) & " - $" & gross_amount(all_income)
                        End If
                    End If
                Next
            End If
            this_month_total_gross = FormatNumber(this_month_total_gross, 2, -1, 0, 0)
            EMWriteScreen this_month_total_gross, 12, 67      'entering the pay information
            updates_to_display = updates_to_display & vbNewLine & "Total income: " & this_month_total_gross
        Else

            For each this_date in this_month_checks_array       'now using the list we made of all the checks for THIS month
                If IsDate(this_date) = TRUE Then
                    the_start_date_to_use = income_start_dt
                    If IsDate(the_start_date_to_use) = False Then the_start_date_to_use = #1/1/1900#
                    If DateDiff("d", this_date, the_start_date_to_use) < 1 Then     'checking to make sure the paydate is not before the income start date - that causes a red line

                        date_found = FALSE          'default for each loop
                        checks_found = 0
                        combined_gross_pay = 0
                        For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
                            If pay_date(all_income) <> "" Then
                                If DateDiff("d", pay_date(all_income), this_date) = 0 Then            'if the pay day matches the date in the list, we add the information to the panel
                                    date_found = TRUE               'saving this so that the information is not over written
                                    checks_found = checks_found + 1
                                    Call create_mainframe_friendly_date(view_pay_date(all_income), jobs_row, 54, "YY")		'pay date
                                    combined_gross_pay = combined_gross_pay + gross_amount(all_income)
                                    total_hours = total_hours + hours(all_income)             'running total of this
                                    updates_to_display = updates_to_display & vbNewLine & "Date - " & view_pay_date(all_income) & " - $" & gross_amount(all_income)
                                    income_items_used = income_items_used & all_income & "~"
                                End If
                            End If
                        Next
                        If date_found = TRUE Then
                            combined_gross_pay = FormatNumber(combined_gross_pay, 2, -1, 0, 0)
                            EMWriteScreen combined_gross_pay, jobs_row, 67              'gross pay - not using the excluded amount on the main panel
                            If checks_found > 1 Then
                                For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
                                    If pay_date(all_income) <> "" Then
                                        If DateDiff("d", pay_date(all_income), this_date) = 0 Then combined_into_one(all_income) = True
                                    End If
                                Next
                            End If
                            jobs_row = jobs_row + 1         'moving to the next row
                        End If

                        If date_found = FALSE Then                  'if the date was not found in the LIST_OF_INCOME_ARRAY - we will use an average
                            Call create_mainframe_friendly_date(this_date, jobs_row, 54, "YY")         'entering the date
                            ave_inc_per_pay = FormatNumber(ave_inc_per_pay, 2, -1, 0, 0)
                            EMWriteScreen ave_inc_per_pay, jobs_row, 67          'entering the average
                            total_hours = total_hours + ave_hrs_per_pay               'totalling hours
                            updates_to_display = updates_to_display & vbNewLine & "Date - " & this_date & " - $" & ave_inc_per_pay
                            jobs_row = jobs_row + 1         'moving to the next row
                        End If
                    End If
                Else
                    testing_run = TRUE
                    script_run_lowdown = script_run_lowdown & vbCR & vbCR & "******* THIS_DATE IS NOT A DATE ********"
                    script_run_lowdown = script_run_lowdown & vbCR & "'this_date' variable is: ~" & this_date & "~"
                    script_run_lowdown = script_run_lowdown & vbCR & "'this_month_checks_array is':"
                    For each dumb_thing in this_month_checks_array
                        script_run_lowdown = script_run_lowdown & vbCR & "~" & dumb_thing & "~"
                    Next
                End If
            Next
            For all_income = 0 to UBound(pay_date)
                If pay_date(all_income) <> "" Then
                    If DatePart("m", view_pay_date(all_income)) = DatePart("m", this_month_checks_array(0)) Then
                        If InStr(income_items_used, "~" & all_income & "~") = 0 Then
                            Call create_mainframe_friendly_date(view_pay_date(all_income), jobs_row, 54, "YY")
                            gross_amount(all_income) = FormatNumber(gross_amount(all_income), 2, -1, 0, 0)
                            EMWriteScreen gross_amount(all_income), jobs_row, 67      'entering the pay information
                            total_hours = total_hours + hours(all_income)             'running total of hours
                            updates_to_display = updates_to_display & vbNewLine & "Date - " & view_pay_date(all_income) & " - $" & gross_amount(all_income)
                            jobs_row = jobs_row + 1
                            income_items_used = income_items_used & all_income & "~"
                        End If
                    End If
                End If
            Next
        End If

        total_hours = Round(total_hours)        'formatting the hours total
        EMWriteScreen "   ", 18, 72             'blanking out BOTH hours positions
        EMWriteScreen "   ", 18, 43
        EMWriteScreen total_hours, 18, 72       'entering the hours on the panel
        updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours
    end sub

	public sub update_panel(update_month)
        'Update the JOBS panel for the month passed to the subroutine

        MAXIS_footer_month = DatePart("m", update_month)        'setting the footer month and year for each month in the list for NAV
        MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
        MAXIS_footer_year = DatePart("yyyy", update_month)
        MAXIS_footer_year = right(MAXIS_footer_year, 2)

        RETRO_month = DateAdd("m", -2, update_month)            'defining 2 months ago for CASH/UNCLE HARRY process
        RETRO_footer_month = DatePart("m", RETRO_month)
        RETRO_footer_month = right("00" & RETRO_footer_month, 2)
        RETRO_footer_year = DatePart("yyyy", RETRO_month)
        RETRO_footer_year = right(RETRO_footer_year, 2)

        'if the active_month is this panel's initial month - then we indicate this month needs to be updated for this panel
        If initial_month_mo = MAXIS_footer_month AND initial_month_yr = MAXIS_footer_year Then update_this_month = TRUE

        'resetting the months arrays for each month loop
        Dim this_month_checks_array(0)
        Dim retro_month_checks_array(0)
        this_month_checks_array(0) = ""
        retro_month_checks_array(0) = ""

        call create_months_check_list(update_month)
        call find_panel

        If update_this_month = TRUE Then              'if this panel should be update in thie month - here is where we do it
            script_run_lowdown = script_run_lowdown & vbCr & "  - Updated JOBS - MEMB " & member & " " & instance & " - " & employer

            months_updated = months_updated & ", " & MAXIS_footer_month & "/" & MAXIS_footer_year       'keeping a list of all the panels updated for each job

            pic_fs_ave_hrs_READ 		= ""
            pic_fs_ave_paycheck_READ 	= ""
            pic_fs_prosp_mo_READ 		= ""
            pic_grh_ave_hrs_READ 		= ""
            pic_grh_ave_paycheck_READ 	= ""
            pic_grh_prosp_mo_READ 		= ""

            If developer_mode = FALSE Then PF9                      'if we are in INQUIRY, the panel is NOT put in edit mode - otherwise here is where it is put in edit mode

            'All of these updates_to_display items are copying what the script is doing in the panel and if in INQUIRY this will be displayed for each job for each month
            updates_to_display = "JOBS Update for " & MAXIS_footer_month & "/" & MAXIS_footer_year & vbNewLine & "MEMBER " & member & " - Employer: " & employer

            'UPDATE Non-check details on the main JOBS panel
            EMWriteScreen left(job_type, 1), 5, 34         'income type, verif and hhourly wage are the samy for each progam
            EMWriteScreen left(verif_type, 1), 6, 34
            EMWriteScreen "      ", 6, 75       'this blanks out the wage otherwise there is carryover and 10.00 becomes 100.00 - which is not correct
            EMWriteScreen hourly_wage, 6, 75
            EMWriteScreen left(pay_freq, 1), 18, 35
            updates_to_display = updates_to_display & vbNewLine & "Income type: " & job_type & " - Verification: " & verif_type & vbNewLine & "Hourly wage: $" & hourly_wage & "/hr. Pay Frequency: " & pay_freq

            If IsDate(income_start_dt) = TRUE Then
                following_job_start = DateAdd("m", 1, income_start_dt)
                Call convert_date_into_MAXIS_footer_month(income_start_dt, job_start_month, job_start_year)
                Call convert_date_into_MAXIS_footer_month(following_job_start, following_job_start_month, following_job_start_year)

                updates_to_display = updates_to_display & vbNewLine & "Income Start Date: " & income_start_dt

                Call write_date(income_start_dt, "MM DD YY", 9, 35)   'entering the start date
            End If


            'Here we update the panel with process and information that is specific to the program the income applies to
            'The order is important here as some will take precedent. Currenly the order iw SNAP-GRH-HC-Cash
            'Since SNAP and GRH budgets are determined by the PIC, the information on the main JOBS panel is less important.
            'Cash has the most specific update requirements for JOBS, so it is last to ensure those are the changes to JOBS that are saved
            'for income that applies to SNAP

            call update_SNAP_PIC

            call update_cash_pic

            call update_main_panel_checks

            call update_hc_estimate

            call update_retro_checks

            If developer_mode = True Then MsgBox updates_to_display
        End If

	end sub

    public sub update_retro_checks()
        'update the RETRO side of the main JOBS panel

        If apply_to_CASH = checked  OR hc_retro = TRUE Then         'now on to cash
            STATS_manualtime = STATS_manualtime + 185
            total_pay = 0
            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "*** Cash Budget Update (or Uncle Harry SNAP) ***"

            update_retro = True
            If DateDiff("m", income_start_dt, RETRO_month) < 0 Then update_retro = False
            If update_retro = False Then updates_to_display = updates_to_display & vbNewLine & "Income Start Date: " & income_start_dt & ", RETRO month: " & RETRO_month & " - CANNOT UPDATE"


            'for each month that is updated for cash, we need to track more detail since there is retro and prosp information to be concerned with
            'this array stores that detail
            If cash_array_info_exists = False Then
                cash_index = 0
                cash_array_info_exists = True
            Else
                cash_index = UBound(cash_info_cash_mo_yr)+1
                ReDim preserve cash_info_cash_mo_yr(cash_index)         'resizing the arrays
                ReDim preserve cash_info_retro_mo_yr(cash_index)
                ReDim preserve cash_info_retro_updtd(cash_index)
                ReDim preserve cash_info_prosp_updtd(cash_index)
                ReDim preserve cash_info_mo_retro_pay(cash_index)
                ReDim preserve cash_info_mo_retro_hrs(cash_index)
                ReDim preserve cash_info_mo_prosp_pay(cash_index)
                ReDim preserve cash_info_mo_prosp_hrs(cash_index)
            End If

            cash_info_retro_updtd(cash_index) = FALSE             'defaults
            cash_info_prosp_updtd(cash_index) = TRUE
            cash_info_cash_mo_yr(cash_index) = MAXIS_footer_month & "/" & MAXIS_footer_year       'save which month we are looking at
            cash_info_retro_mo_yr(cash_index) = RETRO_footer_month & "/" & RETRO_footer_year

            If update_retro = True Then
                If retro_month_checks_array(0) <> "" Then

                    checks_in_month = 0
                    check_to_enter = "~"
                    For all_income = 0 to UBound(pay_date)
                        If pay_date(all_income) <> "" Then
                            If DatePart("m", view_pay_date(all_income)) = DatePart("m", retro_month_checks_array(0)) Then
                                If InStr(check_to_enter, "~"& view_pay_date(all_income) &"~") = 0 Then
                                    checks_in_month = checks_in_month + 1
                                    check_to_enter = check_to_enter & view_pay_date(all_income) & "~"
                                End If
                            End If
                        End If
                    Next

                    If checks_in_month > 5 Then
                        Call convert_date_into_MAXIS_footer_month(retro_month_checks_array(0), retro_footer_month, retro_footer_year)

                        If Instr(mo_w_more_5_chcks, retro_footer_month & "/" & retro_footer_year) = 0 Then
                            mo_w_more_5_chcks = mo_w_more_5_chcks & " " & retro_footer_month & "/" & retro_footer_year
                        End if

                        EMWriteScreen retro_footer_month, 12, 25
                        EMWriteScreen "01", 12, 28
                        EMWriteScreen retro_footer_year, 12, 31
                        updates_to_display = updates_to_display & vbNewLine & "GROUPING ALL " & retro_footer_month & "/" & retro_footer_year & "checks"

                        this_month_total_gross = 0
                        For all_income = 0 to UBound(pay_date)
                            If pay_date(all_income) <> "" Then
                                If DatePart("m", view_pay_date(all_income)) = DatePart("m", retro_month_checks_array(0)) Then
                                    this_month_total_gross = this_month_total_gross + gross_amount(all_income)
                                    total_hours = total_hours + hours(all_income)             'running total of hours
                                    updates_to_display = updates_to_display & vbNewLine & "Date - " & view_pay_date(all_income) & " - $" & gross_amount(all_income)
                                End If
                            End If
                        Next
                        this_month_total_gross = FormatNumber(this_month_total_gross, 2, -1, 0, 0)
                        EMWriteScreen this_month_total_gross, 12, 38      'entering the pay information
                        total_pay = this_month_total_gross
                        updates_to_display = updates_to_display & vbNewLine & "Total income: " & this_month_total_gross
                        count_checks = checks_in_month
                    Else

                        jobs_row = 12           'set for the loop
                        total_hours = 0
                        total_pay = 0
                        count_checks = 0
                        income_items_used = "~"
                        updates_to_display = updates_to_display & vbNewLine & "--- RETRO ---"
                        For each this_date in retro_month_checks_array          'there is a seperate list of retro pay dates
                            If IsDate(this_date) = TRUE Then
                                date_found = FALSE      'default for the start of each loop
                                checks_found = 0
                                combined_gross_pay = 0
                                For all_income = 0 to UBound(pay_date)   'then loop through all of the income information
                                    'conditional if it is the right panel AND the order matches - then do the thing you need to do
                                    If pay_date(all_income) <> "" Then
                                        If DateDiff("d", pay_date(all_income), this_date) = 0 Then        'if the date was in the array, we use it
                                            date_found = TRUE
                                            checks_found = checks_found + 1
                                            cash_info_retro_updtd(cash_index) = TRUE      'setting this to know there was retro information added
                                            Call create_mainframe_friendly_date(view_pay_date(all_income), jobs_row, 25, "YY")		'pay date
                                            combined_gross_pay = combined_gross_pay + gross_amount(all_income)
                                            total_hours = total_hours + hours(all_income)             'running total of hours
                                            total_pay = total_pay + gross_amount(all_income)          'running total of pay for RETRO month only
                                            count_checks = count_checks + 1                 'need to track the number of checks that we used
                                            updates_to_display = updates_to_display & vbNewLine & "Date - " & view_pay_date(all_income) & " - $" & gross_amount(all_income)
                                            income_items_used = income_items_used & all_income & "~"
                                        End If
                                    End If
                                Next
                                If date_found = TRUE Then
                                    combined_gross_pay = FormatNumber(combined_gross_pay, 2, -1, 0, 0)
                                    EMWriteScreen combined_gross_pay, jobs_row, 38              'gross pay - not using the excluded amount on the main panel
                                    If checks_found > 1 Then
                                        For all_income = 0 to UBound(LIST_OF_INCOME_ARRAY, 2)   'then loop through all of the income information
                                            If pay_date(all_income) <> "" Then
                                                If DateDiff("d", pay_date(all_income), this_date) = 0 Then combined_into_one(all_income) = True
                                            End If
                                        Next
                                    End If
                                    jobs_row = jobs_row + 1         'moving to the next row
                                End If
                                'there is no functionality for if pay was not found as we don't use averages on the retro side
                            Else
                                'This is an output for an error report if something went wrong with date understanding
                                testing_run = TRUE
                                script_run_lowdown = script_run_lowdown & vbCR & vbCR & "******* THIS_DATE IS NOT A DATE ********"
                                script_run_lowdown = script_run_lowdown & vbCR & "'this_date' variable is: ~" & this_date & "~"
                                script_run_lowdown = script_run_lowdown & vbCR & "'retro_month_checks_array is':"
                                For each dumb_thing in retro_month_checks_array
                                    script_run_lowdown = script_run_lowdown & vbCR & "~" & dumb_thing & "~"
                                Next
                            End If
                        Next
                        For all_income = 0 to UBound(pay_date)
                            If pay_date(all_income) <> "" Then
                                If retro_month_checks_array(0) <> "" Then
                                    If DatePart("m", view_pay_date(all_income)) = DatePart("m", retro_month_checks_array(0)) Then
                                        If InStr(income_items_used, "~" & all_income & "~") = 0 Then
                                            cash_info_retro_updtd(cash_index) = TRUE      'setting this to know there was retro information added
                                            Call create_mainframe_friendly_date(view_pay_date(all_income), jobs_row, 25, "YY")		'pay date
                                            gross_amount(all_income) = FormatNumber(gross_amount(all_income), 2, -1, 0, 0)
                                            EMWriteScreen gross_amount(all_income), jobs_row, 38      'entering the pay information
                                            total_hours = total_hours + hours(all_income)             'running total of hours
                                            total_pay = total_pay + gross_amount(all_income)          'running total of pay for RETRO month only
                                            count_checks = count_checks + 1                 'need to track the number of checks that we used
                                            updates_to_display = updates_to_display & vbNewLine & "Date - " & view_pay_date(all_income) & " - $" & gross_amount(all_income)
                                            jobs_row = jobs_row + 1
                                            income_items_used = income_items_used & all_income & "~"
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If

                Else
                    updates_to_display = updates_to_display & vbNewLine & " - No retro paydates to update. - "
                    script_run_lowdown = script_run_lowdown & vbNewLine & "--- RETRO ARRAY WAS EMPTY ---"
                End If

                total_hours = Round(total_hours)        'entering retro hours on the retro side
                If total_pay <> 0 Then
                    EMWriteScreen "   ", 18, 43

                    If code_for_six_month_reporting_workaround = True or count_checks <> 0 Then           'if there we check found we make an average of pay and hours for the RETRO side only
                        EMWriteScreen total_hours, 18, 43

                        cash_info_mo_retro_pay(cash_index) = FormatNumber(total_pay, 2,,0)        'save the total hours and pay to the array for CNote
                        cash_info_mo_retro_hrs(cash_index) = total_hours
                        updates_to_display = updates_to_display & vbNewLine & "        Total Hours: " & total_hours
                    End If
                End If

            End If
            EMReadScreen total_hours, 3, 18, 72
            total_hours = trim(total_hours)
            total_hours = replace(total_hours, "_", "")
            If total_hours = "" Then total_hours = 0
            total_hours = total_hours * 1
            cash_info_mo_prosp_hrs(cash_index) = total_hours          'saving the information to the array

            transmit
            EMReadScreen gross_prosp_amount, 8, 17, 67
            gross_prosp_amount = trim(gross_prosp_amount)
            If gross_prosp_amount = "" Then gross_prosp_amount = 0
            cash_info_mo_prosp_pay(cash_index) = FormatNumber(gross_prosp_amount, 2,,0)

            next_cash_month = next_cash_month + 1       'this is for incrementing the array for the next loop
        End If
    end sub

    public sub update_SNAP_PIC()
        'update the SNAP PIC for this JOBS panel

        If apply_to_SNAP = checked Then
            STATS_manualtime = STATS_manualtime + 165
            updates_to_display = updates_to_display & vbNewLine & vbNewLine & "** SNAP Budget Update **" & vbNewLine & "---- PIC ----"
            EMWriteScreen "X", 19, 38       'opening the PIC
            transmit
            PF20

            EMWriteScreen "      ", 8, 64
            EMWriteScreen "        ", 9, 66
            EMWriteScreen "        ", 13, 66
            EMWriteScreen "  ", 14, 64

            list_row = 9                    'here we clear the PIC of all previous data
            beg_of_list_check = ""
            number_of_loops = 0
            Do
                EMWriteScreen "  ", list_row, 13
                EMWriteScreen "  ", list_row, 16
                EMWriteScreen "  ", list_row, 19
                EMWriteScreen "        ", list_row, 25
                EMWriteScreen "      ", list_row, 35
                list_row = list_row + 1
                number_of_loops = number_of_loops + 1

                If list_row = 14 Then
                    transmit
                    PF19

                    EMReadScreen beg_of_list_check, 10, 20, 18
                    list_row = 9
                End If
                If number_of_loops = 30 Then Exit Do
            Loop until beg_of_list_check = "FIRST PAGE"

            Call create_mainframe_friendly_date(date, 5, 34, "YY")                     'enter the current date in date of calculation field
            EMWriteScreen left(pay_freq, 1), 5, 64        'enter the pay frequency code only in the correct field

            'If we are in the footer month that is the first month of income for a new job
            'there is special functionality to LUMP the income in the PIC so that an accurate amount can be entered NEED POLICY REFERENCE
            If job_start_month = MAXIS_footer_month AND job_start_year = MAXIS_footer_year Then
                updates_to_display = updates_to_display & vbNewLine & "Date of Calculation: " & date & "  Pay Frequency: 1 - Monthly"

                fs_appl_date = DateValue(fs_appl_date)      'setting some defaults - more for-nexts are coming
                appl_month_gross = 0
                appl_month_hours = 0

                checks_lumped = ""
                expected_pay_lumped = ""
                income_items_used = "~"

                month_lumped = MAXIS_footer_month & "/" & MAXIS_footer_year     'formatting for readability
                reason_lumped = "first month of new job"
                For each this_date in this_month_checks_array           'this array was set at the begining of this month's loop - it will get us all our pay dates
                    If IsDate(this_date) = TRUE Then
                        the_start_date_to_use = income_start_dt
                        If IsDate(the_start_date_to_use) = False Then the_start_date_to_use = #1/1/1900#
                        If DateDiff("d", this_date, the_start_date_to_use) < 1 Then     'if the pay date we are looking at is on or after the income start date we will add it in to the lump

                            date_found = FALSE      'default for before we look at the checks
                            For all_income = 0 to UBound(pay_date)                  'then loop through all of the income information
                                If DateDiff("d", pay_date(all_income), this_date) = 0 Then            'if the pay date (this is NOT the view_pay_date) then we use the amount provided on ENTER PAY
                                    date_found = TRUE

                                    appl_month_gross = appl_month_gross + gross_amount(all_income)
                                    appl_month_hours = appl_month_hours + hours(all_income)
                                    checks_lumped = checks_lumped & view_pay_date(all_income) & " - $" & gross_amount(all_income) & " - " & hours(all_income) & "hrs.; "      'saving a list of the checks used
                                    income_items_used = income_items_used & all_income & "~"
                                End If
                            Next

                            If date_found = FALSE Then          'if the date was not provided on ENTER PAY, we are going to use an averate
                                appl_month_gross = appl_month_gross + ave_inc_per_pay
                                appl_month_hours = appl_month_hours + ave_hrs_per_pay

                                expected_pay_lumped = expected_pay_lumped & this_date & " - anticipated $" & ave_inc_per_pay & " - " & ave_hrs_per_pay & "hrs.; "        'saving a list of the pay estimates used
                            End If
                        End If
                    Else
                        testing_run = TRUE
                        script_run_lowdown = script_run_lowdown & vbCR & vbCR & "******* THIS_DATE IS NOT A DATE ********"
                        script_run_lowdown = script_run_lowdown & vbCR & "'this_date' variable is: ~" & this_date & "~"
                        script_run_lowdown = script_run_lowdown & vbCR & "'this_month_checks_array is':"
                        For each dumb_thing in this_month_checks_array
                            script_run_lowdown = script_run_lowdown & vbCR & "~" & dumb_thing & "~"
                        Next
                    End If
                Next
                For all_income = 0 to UBound(pay_date)
                    If DatePart("m", view_pay_date(all_income)) = DatePart("m", this_month_checks_array(0)) Then
                        If InStr(income_items_used, "~" & all_income & "~") = 0 Then
                            appl_month_gross = appl_month_gross + gross_amount(all_income)
                            appl_month_hours = appl_month_hours + hours(all_income)
						    If IsNumeric(exclude_SNAP_hours(all_income)) Then appl_month_hours = appl_month_hours - exclude_SNAP_hours(all_income)

                            checks_lumped = checks_lumped & view_pay_date(all_income) & " - $ " & gross_amount(all_income) & " - " & hours(all_income) & "hrs.; "      'saving a list of the checks used
                            income_items_used = income_items_used & all_income & "~"
                        End If
                    End If
                Next

                EMWriteScreen "1", 5, 64                        'A pay frequency of 1 (monthly) is entered on the PIC
                EMWriteScreen MAXIS_footer_month, 9, 13         'a default date of the first of the month is entered on the PIC
                EMWriteScreen "01", 9, 16
                EMWriteScreen MAXIS_footer_year, 9, 19
                appl_month_gross = FormatNumber(appl_month_gross, 2, -1, 0, 0)      'the sum of all the pay check gross amounts (or average) is formated and entered
                EMWriteScreen appl_month_gross, 9, 25
                appl_month_hours = FormatNumber(appl_month_hours, 2, -1, 0, 0)
                EMWriteScreen appl_month_hours, 9, 35
                updates_to_display = updates_to_display & vbNewLine & "Actual Pay: Date - " & MAXIS_footer_month & "/01/" & MAXIS_footer_year & " - $" & appl_month_gross & " - " & appl_month_hours & " hrs." & vbNewLine

                If checks_lumped <> "" Then     'formatting the lists of the checks that we included for the CNote
                    If right(checks_lumped, 2) = "; " Then checks_lumped = left(checks_lumped, len(checks_lumped)-2)
                End If
                If expected_pay_lumped <> "" Then
                    If right(expected_pay_lumped, 1) = "; " Then expected_pay_lumped = left(expected_pay_lumped, len(expected_pay_lumped)-2)
                End If

                income_lumped_mo  = month_lumped              'saving all the information about the lumping to the array for CNoting
                lump_reason       = reason_lumped
                act_checks_lumped = checks_lumped
                est_checks_lumped = expected_pay_lumped
                lump_gross        = appl_month_gross
                lump_hrs          = appl_month_hours

            Else        'this is if we are in any month other than the month of application or first month of income for this job
                updates_to_display = updates_to_display & vbNewLine & "Date of Calculation: " & date & "  Pay Frequency: " & pay_freq
                If pick_one = use_estimate Then           'if use estimate was selected - then we just plug in the hours/wk and pay/hr
                    updates_to_display = updates_to_display & vbNewLine & "Estimate: Hourly Wage - $" & hourly_wage & "/hr - " & snap_hrs_per_wk & " hrs/wk" & vbNewLine
                    EMWriteScreen hourly_wage, 9, 66
                    the_hrs_per_wk = FormatNumber(snap_hrs_per_wk, 2, -1, 0, 0)
                    EMWriteScreen snap_hrs_per_wk, 8, 64
                End If
                If pick_one = use_actual Then             'if we have to use actual - it is harder - HERE WE GO!
                    updates_to_display = updates_to_display & vbNewLine & "Actual Pay: Date - "

                    'here we do not need to compare it to expected checks because the PIC dates do not need to align with the current month
                    list_row = 9
                    For order_number = 1 to order_ubound                        'loop through the order number lowest to highest
                        For all_income = 0 to UBound(pay_date)                  'then loop through all of the income information
                            If check_order(all_income) = order_number Then
                                If NOT exclude_entirely(all_income) and exclude_from_SNAP(all_income) = unchecked  Then
                                    If MX_region <> "INQUIRY DB" Then
                                        If list_row = 14 Then
                                            Do
                                                EMWaitReady 0, 0
                                                PF20
                                                EMWaitReady 0, 0
                                                EMReadScreen top_check_blank, 8, 9, 25
                                            Loop until top_check_blank = "________"
                                            list_row = 9
                                        End If
                                    End If
                                    Call create_mainframe_friendly_date(view_pay_date(all_income), list_row, 13, "YY")           'this is the CIEW date - the one the owrker actually entered
                                    calc_exclude = 0
                                    If IsNumeric(exclude_ALL_amount(all_income)) Then calc_exclude = calc_exclude + exclude_ALL_amount(all_income)
                                    If IsNumeric(split_exclude_amount(all_income)) Then calc_exclude = calc_exclude + split_exclude_amount(all_income)
                                    If IsNumeric(exclude_SNAP_amount(all_income)) Then calc_exclude = calc_exclude + exclude_SNAP_amount(all_income)
                                    net_amount = gross_amount(all_income) - calc_exclude      'taking out excluded amounts
                                    net_amount = FormatNumber(net_amount, 2, -1, 0, 0)
						            the_hours = hours(all_income)
                                    If IsNumeric(exclude_SNAP_hours(all_income)) Then the_hours = the_hours - exclude_SNAP_hours(all_income)
        							If pay_excld_ot(all_income) = checked and IsNumeric(pay_split_ot_hours(all_income)) Then the_hours = the_hours - pay_split_ot_hours(all_income)

                                    the_hours = FormatNumber(the_hours, 2, -1, 0, 0)
                                    EMWriteScreen net_amount, list_row, 25      'entering the pay amount to count
                                    EMWriteScreen the_hours, list_row, 35     'enting the hours

                                    updates_to_display = updates_to_display & view_pay_date(all_income) & " - $" & net_amount & " - " & hours(all_income) & " hrs." & vbNewLine
                                    list_row = list_row + 1         'next line of the PIC'

                                End If
                            End If
                        next
                    next

                End If
            End If
            read_calculation = True
            Do
                transmit            'saving the PIC
                EMReadScreen escape_route, 26, 20, 6
                If read_calculation = True and escape_route <> "WARNING: ENTER CASE NOTE D" Then
                    read_calculation = False

                    EMReadScreen pic_fs_ave_hrs_READ, 		8, 16, 52
                    EMReadScreen pic_fs_ave_paycheck_READ, 	10, 17, 54
                    EMReadScreen pic_fs_prosp_mo_READ, 		10, 18, 54

                    pic_fs_ave_hrs_READ = trim(pic_fs_ave_hrs_READ)
                    pic_fs_ave_paycheck_READ = trim(pic_fs_ave_paycheck_READ)
                    pic_fs_prosp_mo_READ = trim(pic_fs_prosp_mo_READ)
                End If
                If escape_route = "CHANGE DATA OR PF3 TO EXIT" Then PF3
                EMReadScreen pic_menu, 43, 3, 22
            Loop until pic_menu <> "Food Support Prospective Income Calculation"
        End If          'If EARNED_INCOME_PANELS_ARRAY(apply_to_SNAP, ei_panel) = checked THen'
    end sub

	public sub ytd_calculator(err_msg)
        'used to determine the amount of checks based on the calculation of YTD pay and hours

		If missing_checks_list = "" Then
			err_msg = err_msg & vbCr & "* You selected to have the script complete a YTD calculation, there does not appear to be a pay date that is missing with the entered checks. Complete the dialog will all of the known checks before running this YTD calcuator."
		Else
			If left(missing_checks_list, 1) = "~" Then missing_checks_list = right(missing_checks_list, len(missing_checks_list) - 1)       'create an array of the missing checks
			missing_checks_array = split(missing_checks_list, "~")
			loop_to_add_missing_checks = TRUE       'these are set to show the ENTER PAY Dialog again without going to he CONFIRM BUDGET Dialog
			review_small_dlg = TRUE

			For each check_missed in missing_checks_array        'these missing dates get added to the LIST_OF_INCOME_ARRAY automatically
				check_date_before = ""
				check_before_index = ""
				check_date_after = ""
				check_after_index = ""
				before_check_ytd_pay = ""
				before_check_ytd_hrs = ""
				after_check_ytd_pay = ""
				after_check_ytd_hrs = ""
				missing_check_rate_of_pay = pay_per_hr

				pay_item = UBound(pay_date)
				If pay_date(pay_item) <> "" Then pay_item = pay_item + 1
				Call resize_check_list(pay_item)
				pay_date(pay_item) = check_missed
				pay_detail_btn(pay_item) = 2000+pay_item
				duplct_pay_date(check_count) = False
				calculated_by_ytd(pay_item) = True
				duplicate_pay_date(pay_item) = False
				check_info_entered(pay_item) = True
				future_check(pay_item) = False
				If DatePart("m", fs_appl_date) = DatePart("m", pay_date(pay_item)) AND DatePart("yyyy", fs_appl_date) = DatePart("yyyy", pay_date(pay_item)) Then   'if the paydate is in the application month
					If DateDiff("d", date, pay_date(pay_item)) > 0 Then future_check(pay_item) = TRUE   'this is a future check
				End If

				For each_expctd_chk = 0 to UBound(expected_check_array)
					If expected_check_array(each_expctd_chk) = check_missed Then
						check_date_before = expected_check_array(each_expctd_chk-1)
						check_date_after = expected_check_array(each_expctd_chk+1)
					End If
				Next

				For each_known_chk = 0 to UBound(pay_date)
					If DateDiff("d", pay_date(each_known_chk), check_date_before) = 0 Then
						check_before_index = each_known_chk
					End If
					If DateDiff("d", pay_date(each_known_chk), check_date_after) = 0 Then
						check_after_index = each_known_chk
					End If
				Next

				Do
					Do
						ytd_err_msg = ""
						cancel_ytd_calc = False
						Dialog1 = ""
						BeginDialog Dialog1, 0, 0, 281, 225, "YTD Check Calculator"
							GroupBox 5, 10, 265, 65, "Missing Check Information"
							Text 20, 25, 135, 10, "Check Date: " & pay_date(pay_item)
							Text 15, 40, 115, 10, "Gross Amount: $ " & gross_amount(pay_item)
							Text 20, 55, 100, 10, " Total Hours: " & hours(pay_item)
							GroupBox 5, 85, 125, 80, "Check Before"
							Text 10, 105, 110, 10, "Check Date: " & pay_date(check_before_index)
							Text 20, 125, 35, 10, "YTD Pay:"
							EditBox 55, 120, 50, 15, before_check_ytd_pay
							Text 15, 145, 40, 10, "YTD Hours:"
							EditBox 55, 140, 50, 15, before_check_ytd_hrs
							GroupBox 140, 85, 130, 105, "Check After"
							Text 145, 105, 110, 10, "Check Date: " & pay_date(check_after_index)
							Text 155, 125, 35, 10, "YTD Pay:"
							EditBox 190, 120, 50, 15, after_check_ytd_pay
							Text 150, 145, 40, 10, "YTD Hours:"
							EditBox 190, 140, 50, 15, after_check_ytd_hrs
							Text 150, 165, 95, 10, "Check Gross Pay: " & gross_amount(check_after_index)
							Text 145, 175, 95, 10, "Check Gross Hours: " & hours(check_after_index)
							Text 5, 170, 130, 10, "You must enter both YTD Pay amounts."
							Text 5, 180, 130, 20, "You must enter both YTD Hours amounts OR list the rate of pay ($/hr)."
							Text 5, 210, 50, 10, "Rate of Pay: $"
							EditBox 60, 205, 50, 15, missing_check_rate_of_pay
							ButtonGroup ButtonPressed
								PushButton 175, 20, 90, 15, "Calculate", calculate_ytd_btn
								PushButton 165, 205, 50, 15, "Done", done_ytd_btn
								CancelButton 220, 205, 50, 15
						EndDialog

						dialog Dialog1
						save_your_work

						If IsNumeric(before_check_ytd_pay) = False Then ytd_err_msg = ytd_err_msg & vbCr & "* The YTD Pay amount from the " & pay_date(check_before_index) & " check must be entered as a number"
						If IsNumeric(after_check_ytd_pay) = False Then ytd_err_msg = ytd_err_msg & vbCr & "* The YTD Pay amount from the " & pay_date(check_after_index) & " check must be entered as a number"
						If IsNumeric(before_check_ytd_hrs) = False or IsNumeric(after_check_ytd_hrs) = False Then
							If IsNumeric(missing_check_rate_of_pay) = False Then ytd_err_msg = ytd_err_msg & vbCr & "* The Rate of Pay OR BOTH YTD Hours must be entered as a number."
						End If

						If ButtonPressed = 0 then
							check_date_before = ""
							check_before_index = ""
							check_date_after = ""
							check_after_index = ""
							before_check_ytd_pay = ""
							before_check_ytd_hrs = ""
							after_check_ytd_pay = ""
							after_check_ytd_hrs = ""

							ytd_err_msg = ""
							cancel_ytd_calc = True

							Call delete_one_check(pay_item)

							ButtonPressed = done_ytd_btn
						End If
						If ytd_err_msg <> "" Then MsgBox ytd_err_msg
						If ButtonPressed = -1 Then ButtonPressed = done_ytd_btn

					Loop until ytd_err_msg = ""
					save_your_work

					If cancel_ytd_calc = False Then
						missing_check_ytd = after_check_ytd_pay - gross_amount(check_after_index)
						gross_amount(pay_item) = missing_check_ytd - before_check_ytd_pay

						ytd_calc_notes(pay_item) = pay_date(pay_item) & " Check amount Calculation: ; "
						ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & pay_date(check_after_index) & " check - YTD: $ " & after_check_ytd_pay & ", Gross Pay: $ " & gross_amount(check_after_index) & "; "
						ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & pay_date(check_before_index) & " check - YTD: $ " & before_check_ytd_pay & "; "
						ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & "$ " & after_check_ytd_pay & " - $ " & gross_amount(check_after_index) & " - $ " & before_check_ytd_pay & " = $ " & gross_amount(pay_item) & "; "
						ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & pay_date(pay_item) & " gross pay: $ " & gross_amount(pay_item) & "; "

						missing_hrs_by_ytd = ""
						missing_hrs_by_rate_of_pay = ""
						hours(pay_item) = ""
						If IsNumeric(before_check_ytd_hrs) = True and IsNumeric(after_check_ytd_hrs) = True Then
							missing_hrs_by_ytd = after_check_ytd_hrs - hours(check_after_index) - before_check_ytd_hrs
						End If
						If IsNumeric(missing_check_rate_of_pay) = True Then
							missing_hrs_by_rate_of_pay = gross_amount(pay_item) / missing_check_rate_of_pay
							ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & "Rate of Pay: $ " & missing_check_rate_of_pay & "/hr. " & pay_date(pay_item) & " Gross pay of $ " & gross_amount(pay_item) & "/" & missing_hrs_by_rate_of_pay & " = " & missing_hrs_by_rate_of_pay & "; "
						End If

						If missing_hrs_by_ytd = missing_hrs_by_rate_of_pay Then
							hours(pay_item)= missing_hrs_by_rate_of_pay
						ElseIf missing_hrs_by_ytd = "" Then
							hours(pay_item) = missing_hrs_by_rate_of_pay
						ElseIf missing_hrs_by_rate_of_pay = "" Then
							hours(pay_item) = missing_hrs_by_ytd
							ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & pay_date(check_after_index) & " check - YTD Hours: " & after_check_ytd_hrs & ", Hours:  " & hours(check_after_index) & "; "
							ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & pay_date(check_before_index) & " check - YTD Hours: " & before_check_ytd_hrs & "; "
							ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & after_check_ytd_hrs & " - " & hours(check_after_index) & " - " & before_check_ytd_hrs & " = " & hours(pay_item) & "; "
						Else
							MsgBox "It appears that we have 2 different ways to calculate the number of hours on the missing check and these processes do not match. This script has calcualted the hours in the following ways: " & vbCr & vbCr &_
									"Based on YTD Hours reported: " & missing_hrs_by_ytd & vbCr &_
									"   " & pay_date(check_after_index) & " check YTD Hours: " & after_check_ytd_hrs & vbCr &_
									" - " & pay_date(check_after_index) & " check total Hours: " & hours(check_after_index) & vbCr &_
									" - " & pay_date(check_before_index) & " check YTD Hours: " & before_check_ytd_hrs & vbCr &_
									" = " & missing_hrs_by_ytd & " hours" & vbCr & vbCr &_
									"Based on Rate of Pay reported: " & missing_hrs_by_rate_of_pay & vbCr &_
									"   " & pay_date(pay_item) & " gross pay: $ " & gross_amount(pay_item) & vbCr &_
									" / " & "Rate of pay: " & missing_check_rate_of_pay & vbCr &_
									" = " & missing_hrs_by_rate_of_pay & " hours" & vbCr & vbCr &_
									"The script will display the YTD Calculator again, either remove the YTD Hours information or Rate of Pay information, whichever is creating an incrorrect hours calculation."
							ButtonPressed = calculate_ytd_btn
						End If
						ytd_calc_notes(pay_item) = ytd_calc_notes(pay_item) & pay_date(pay_item) & " total hours: " & hours(pay_item)

						pay_date(pay_item) = DateAdd("d", 0, pay_date(pay_item))
						gross_amount(pay_item) = FormatNumber(gross_amount(pay_item), 2, -1, 0, 0)
					End If
					save_your_work
				Loop until ButtonPressed = done_ytd_btn
			Next
		End If

	end sub


end class

Dim JOBS_PANELS()
ReDim JOBS_PANELS(0)

'some script wide variables
Dim fs_appl_date, fs_appl_footer_month, fs_appl_footer_year, panel_exists, panel_created
Dim fs_prog_status, cash_one_status, cash_two_status, grh_status, snap_status, hc_status
Dim CASH_case, SNAP_case, HC_case, GRH_case, cash_one_prog, cash_two_prog
Dim HH_member_array()


the_first_of_CM_2 = CM_plus_2_mo & "/1/" & CM_plus_2_yr     'this is setting start and end dates for creating a list
CM_2_mo = DatePart("m", the_first_of_CM_2)
CM_2_yr = DatePart("yyyy", the_first_of_CM_2)


'Constants to make an option selection easier to read.
const use_actual        = 1
const use_estimate      = 2

'FUNCTION ==================================================================================================================

function offer_new_panel_creation()
	Do
		'THIS VARIABLE SET TO FALSE WILL CAUSE THE SCRIPT TO END AFTER THIS LOOP
		y_pos = 25              'setting coordinates for the dialog to be created - this is the vertical position in the dialog
		dlg_len = 15 * UBOUND(JOBS_PANELS) + 15 * (Round(UBOUND(HH_member_array) / 4)+1) + 125       'creating the height of the dialog

		'ASK TO ADD NEW PANEL Dailog - lists all current panels, Yes/No question about adding another
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 390, dlg_len, "Do you want to add a new JOBS or BUSI Panel?"

			Text 5, 10, 105, 10, "Known JOBS and BUSI panels:"        'This part lists the current panels and will change each time through the loop as new panels are added '

			'Looping through the JOBS_PANELS to get all panels found earlier
			If panel_exists = True Then
				For panel = 0 to UBOUND(JOBS_PANELS)
					'compiling the information about the panel here to make it more readble and specific to the PANEL detail to include
					earned_income_panel_detail = "JOBS " & JOBS_PANELS(panel).member & " " & JOBS_PANELS(panel).instance & " - " & JOBS_PANELS(panel).employer
					earned_income_panel_detail = earned_income_panel_detail & " - Income Start: " & JOBS_PANELS(panel).income_start_dt & " - Verif: " & JOBS_PANELS(panel).old_verif

					Text 10, y_pos, 375, 10, earned_income_panel_detail     'here is where we actually list the information about the panel on the dialog
					y_pos = y_pos + 15                                      'incrementing this placeholder so that panel information is not on top of each other in the dialog
				Next
			Else
				Text 10, y_pos, 375, 10, "** THERE ARE NO EARNED INCOME PANELS ON THIS CASE **"     'here is where we actually list the information about the panel on the dialog
				y_pos = y_pos + 15                                      'incrementing this placeholder so that panel information is not on top of each other in the dialog
			End If
			y_pos = y_pos + 5     'now we move down a little more in the dialog
			Text 5, y_pos, 295, 10, "These are all the panels that are currently known in MAXIS for these Household Members:" 'listing all the household members we looked at in gathering panel information
			y_pos = y_pos + 15
            x_pos = 20
			For each member in HH_member_array
				Text x_pos, y_pos, 45, 10, "MEMB " & member
                x_pos = x_pos + 75
                If x_pos > 250 Then
                    x_pos = 20
                    y_pos = y_pos + 15
                End If
				' y_pos = y_pos + 15
				'Text 10, 75, 45, 10, "MEMBER 01"
			Next
            If ((UBOUND(HH_member_array)+1) Mod 4) <> 0 Then y_pos = y_pos + 15
			' y_pos = y_pos - 10
			' Text 80, y_pos, 160, 10, "Do you need to add a new JOBS or BUSI panel?"     'FUTURE FUNCTIONALITY - Saving this for when BUSI is added
			Text 115, y_pos, 160, 10, "Do you need to add a new JOBS panel?"
			ButtonGroup ButtonPressed             'NO CANCEL button on this dialog - pressing the 'X' in the corner will default to the 'No' button pressed
				PushButton 120, y_pos + 15, 140, 20, "Yes - Add a new Earned Income panel", add_new_panel_button
				PushButton 120, y_pos + 35, 140, 20, "No - The panel(s) to update are in MAXIS", continue_to_update_button
		EndDialog

		dialog Dialog1          'displaying this dialog - no error handling because no input fields - just a Yes/No
								'vbYesNo MsgBox NOT used because we want to allow MAXIS navigation in this time.

		'Pushing the 'Yes' button on the dialog causes this code to be used - pressing 'No' will skip this
		If buttonpressed = add_new_panel_button Then

			original_month = MAXIS_footer_month     'saving the MAXIS footer month and year because we may move around as the panels are added in the month the income started OR application month
			original_year = MAXIS_footer_year
			panel_to_add = "JOBS"                   'defaulting this to JOBS because that is actually the only option right now
			add_panel_button_pushed_count = add_panel_button_pushed_count + 1

			If panel_to_add = "" Then
				'TYPE OF PANEL TO ADD Dialog - Select panel type
				'FUTURE FUNCTIONALITY - add BUSI back as an option to select here
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 191, 50, "Panel Type to Add"
					' DropListBox 30, 30, 60, 45, "Select one..."+chr(9)+"JOBS"+chr(9)+"BUSI", panel_to_add
					DropListBox 30, 30, 60, 45, "Select one..."+chr(9)+"JOBS", panel_to_add
					ButtonGroup ButtonPressed
						OkButton 135, 10, 50, 15
						CancelButton 135, 30, 50, 15
					Text 15, 10, 85, 20, "Which type of panel would you like to add?"
				EndDialog

				cancel_clarify = ""     'resetting this here - this supports canceling the new job add without cancelling the script
				Do
					Do
						err_msg = ""

						dialog Dialog1

						If ButtonPressed = 0 then       'this is the cancel button
							'this is an upgrade in functionality from cancel_confirmation to asking if we should cancel the current function or the script
							cancel_clarify = MsgBox("Do you want to stop the script entirely?" & vbNewLine & vbNewLine & "If the script is stopped no actions taken so far will be noted.", vbQuestion + vbYesNo, "Clarify Cancel")
							If cancel_clarify = vbYes Then script_end_procedure("~PT: user pressed cancel")
							'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
						End if
						If cancel_clarify = vbNo Then       'This means to cancel the current operation but not the script
							panel_to_add = ""               'blanking this out so a panel is not added
							Exit Do                         'leaving the loop for the dialog
						End If

						If panel_to_add = "Select one..." Then err_msg = err_msg & vbNewLine & "* Indicate which type of panel needs to be added."      'error handling - must select a panel type

						If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg                                                'showing the error message

					Loop until err_msg = ""
					call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
				LOOP UNTIL are_we_passworded_out = false
			End If

			If panel_to_add = "JOBS" Then
				ReDim Preserve JOBS_PANELS(panels_count)				'resizing the array
				Set JOBS_PANELS(panels_count) = new jobs_income

				If CASH_case = TRUE Then JOBS_PANELS(panels_count).apply_to_CASH = checked     'These are defaulted by whatever program is active or pending - will be able to be changed later
				If SNAP_case = TRUE Then JOBS_PANELS(panels_count).apply_to_SNAP = checked
				If HC_case = TRUE Then JOBS_PANELS(panels_count).apply_to_HC = checked
				If GRH_case = TRUE Then JOBS_PANELS(panels_count).apply_to_GRH = checked

                panel_created = False
                JOBS_PANELS(panels_count).create_new_panel()			'capture new panel information

				If panel_created Then
					panels_count = panels_count + 1       'incrementing our counter to be ready for the next panel/member/income type
                    panel_exists = True
				Else
					ReDim Preserve JOBS_PANELS(panels_count-1)				'resizing the array
				End If
			End If

			MAXIS_footer_month = original_month     'resetting the footer month and year to what was indicated in the initial dialog
			MAXIS_footer_year = original_year
			Call back_to_SELF
		End If      'If buttonpressed = add_new_panel_button Then
		'There is nothing specific that happens if the 'NO' or continue_to_update button is pushed other than leaving this portion of the functionality to the next portion

		'this loop until functionality allows for as many JOBS/BUSI panels to be added as needed.
		'Also, note that the new panel will be in the array and so will be added to the dialog asking about adding a new panel
	Loop until ButtonPressed = continue_to_update_button
end function

function read_txt_value(variable_name, text_info, variable_type)
    variable_type = UCase(variable_type)
    Select Case variable_type
        Case "DATE"
            variable_name = text_info
            If IsDate(variable_name) Then variable_name = DateAdd("d", 0, variable_name)
        Case "NUMBER"
            variable_name = text_info
            If IsNumeric(variable_name) Then variable_name = variable_name * 1
        Case "INTEGER"
            variable_name = text_info
            If IsNumeric(variable_name) Then variable_name = Round(variable_name)
        Case "AMOUNT"
            variable_name = text_info
            If IsNumeric(variable_name) then variable_name = FormatNumber(variable_name, 2, -1, 0, 0)
        Case "BOOLEAN"
            variable_name = ""
            If text_info = "" Then variable_name = False
            If UCase(text_info) = "TRUE" Then variable_name = True
            If UCase(text_info) = "FALSE" Then variable_name = False
        Case "CHECKBOX"
            If IsNumeric(text_info) Then
                If text_info = 1 Then variable_name = checked
                If text_info = 0 Then variable_name = unchecked
            End If
            If text_info = "" Then variable_name = unchecked
            If text_info = "0" Then variable_name = unchecked
            If text_info = "1" Then variable_name = checked
        Case Else
            variable_name = text_info
    End Select

end function

function read_txt_array(variable_name, text_info, variable_type, delimiter, need_redim)
    temp_array = split(text_info, delimiter)
    If need_redim Then ReDim variable_name(UBound(temp_array))
    For the_thing = 0 to UBound(temp_array)
        variable_name(the_thing) = temp_array(the_thing)
        Call read_txt_value(variable_name(the_thing), temp_array(the_thing), variable_type)
    Next
end function

function restore_your_work(vars_filled)
'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run

	'Now determines name of file
	local_changelog_path = user_myDocs_folder & "earned-income-detail-" & MAXIS_case_number & "-info.txt"

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then

			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

			If pull_variables = vbYes Then
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(local_changelog_path, ForReading)

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				saved_eib_details = split(every_line_in_text_file, vbNewLine)
				vars_filled = TRUE

				panel_count = 0
				check_count = 0
                filling_class_variables = False
                job_count = ""

				For Each text_line in saved_eib_details
					If trim(text_line) <> "" Then
						var_array = ""
						var_array = split(text_line, "~%~%~%~%~")

                        If NOT filling_class_variables Then
                            If var_array(0) = "fs_prog_status"      Then Call read_txt_value(fs_prog_status, var_array(1), "String")
                            If var_array(0) = "fs_appl_date"        Then Call read_txt_value(fs_appl_date, var_array(1), "Date")
                            If var_array(0) = "fs_appl_footer_month" Then Call read_txt_value(fs_appl_footer_month, var_array(1), "String")
                            If var_array(0) = "fs_appl_footer_year" Then Call read_txt_value(fs_appl_footer_year, var_array(1), "String")
                            If var_array(0) = "cash_one_status"     Then Call read_txt_value(cash_one_status, var_array(1), "String")
                            If var_array(0) = "cash_two_status"     Then Call read_txt_value(cash_two_status, var_array(1), "String")
                            If var_array(0) = "grh_status"          Then Call read_txt_value(grh_status, var_array(1), "String")
                            If var_array(0) = "snap_status"         Then Call read_txt_value(snap_status, var_array(1), "String")
                            If var_array(0) = "hc_status"           Then Call read_txt_value(hc_status, var_array(1), "String")
                            If var_array(0) = "CASH_case"           Then Call read_txt_value(CASH_case, var_array(1), "Boolean")
                            If var_array(0) = "SNAP_case"           Then Call read_txt_value(SNAP_case, var_array(1), "Boolean")
                            If var_array(0) = "HC_case"             Then Call read_txt_value(HC_case, var_array(1), "Boolean")
                            If var_array(0) = "GRH_case"            Then Call read_txt_value(GRH_case, var_array(1), "Boolean")
                            If var_array(0) = "cash_one_prog"       Then Call read_txt_value(cash_one_prog, var_array(1), "String")
                            If var_array(0) = "cash_two_prog"       Then Call read_txt_value(cash_two_prog, var_array(1), "String")
                            If var_array(0) = "panel_exists"        Then Call read_txt_value(panel_exists, var_array(1), "Boolean")
                            If var_array(0) = "panel_created"       Then Call read_txt_value(panel_created, var_array(1), "Boolean")

                            If var_array(0) = "HH_member_array" Then Call read_txt_array(HH_member_array, var_array(1), variable_type, "|", True)

                            If var_array(0) = "CLASS_INFO" Then
                                numb_of_jobs_classes = var_array(1)
                                If NOT IsNumeric(numb_of_jobs_classes) Then numb_of_jobs_classes = 0
                                numb_of_jobs_classes = numb_of_jobs_classes * 1
                                filling_class_variables = True

                                ReDim JOBS_PANELS(numb_of_jobs_classes)									'resizing the array
                                job_count = 0
                            End If

                        Else
                            If var_array(0) = "SET_NEW" Then
                                Set JOBS_PANELS(job_count) = new jobs_income
                                JOBS_PANELS(job_count).next_check_btn = 150
                                JOBS_PANELS(job_count).cancel_check_btn = 160
                                JOBS_PANELS(job_count).save_details_btn = 170
                                JOBS_PANELS(job_count).delete_check_btn = 180
                            ElseIf var_array(0) = "NEXT_CLASS" Then
                                ' filling_class_variables = False
                                ' job_count = ""
                                job_count = job_count + 1
                            Else
                                If job_count <> "" Then
                                    call JOBS_PANELS(job_count).restore_info(var_array(0), var_array(1))
                                End If
                            End If
                        End If

					End If
				Next
			End If
		End If
	End With

end function

function save_your_work()
'This function records the variables into a txt file so that it can be retrieved by the script if run later.

	'Now determines name of file
	local_changelog_path = user_myDocs_folder & "earned-income-detail-" & MAXIS_case_number & "-info.txt"

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then
			.DeleteFile(local_changelog_path)
		End If

		'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

		If .FileExists(local_changelog_path) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(local_changelog_path, ForWriting, true)

			objTextStream.WriteLine "fs_prog_status~%~%~%~%~" & fs_prog_status
			objTextStream.WriteLine "fs_appl_date~%~%~%~%~" & fs_appl_date
			objTextStream.WriteLine "fs_appl_footer_month~%~%~%~%~" & fs_appl_footer_month
			objTextStream.WriteLine "fs_appl_footer_year~%~%~%~%~" & fs_appl_footer_year
			objTextStream.WriteLine "cash_one_status~%~%~%~%~" & cash_one_status
			objTextStream.WriteLine "cash_two_status~%~%~%~%~" & cash_two_status
			objTextStream.WriteLine "grh_status~%~%~%~%~" & grh_status
			objTextStream.WriteLine "snap_status~%~%~%~%~" & snap_status
			objTextStream.WriteLine "hc_status~%~%~%~%~" & hc_status
			objTextStream.WriteLine "CASH_case~%~%~%~%~" & CASH_case	'BOOL
			objTextStream.WriteLine "SNAP_case~%~%~%~%~" & SNAP_case	'BOOL
			objTextStream.WriteLine "HC_case~%~%~%~%~" & HC_case		'BOOL
			objTextStream.WriteLine "GRH_case~%~%~%~%~" & GRH_case	    'BOOL
			objTextStream.WriteLine "cash_one_prog~%~%~%~%~" & cash_one_prog
			objTextStream.WriteLine "cash_two_prog~%~%~%~%~" & cash_two_prog
			objTextStream.WriteLine "panel_exists~%~%~%~%~" & panel_exists	    'BOOL
			objTextStream.WriteLine "panel_created~%~%~%~%~" & panel_created	'BOOL

            objTextStream.WriteLine "HH_member_array~%~%~%~%~" & join(HH_member_array, "|")   'array of household members considered in the script

            objTextStream.WriteLine "CLASS_INFO~%~%~%~%~" & UBound(JOBS_PANELS)
            For jb_panel = 0 to UBOUND(JOBS_PANELS)       'looping through all of the current JOBS or BUSI panels
                JOBS_PANELS(jb_panel).save_info(objTextStream)
            Next
        End If
    End With
end function

'USER INFORMATION ==========================================================================================================
'in the list checks dialog
pay_freq_msg_text = "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
					"ABOUT ENTERING PAY FREQUENCY"  & vbNewLine & vbNewLine &_
					"The script does not fill this field with the pay frequency that may be entered on the panel. This field needs to be avaluated and updated every time the script is run."  & vbNewLine & vbNewLine &_
					"PAY FREQUENCY CODING ERRORS HAVE A HIGH RATE OF PAYMENT ACCURACY ERRORS."  & vbNewLine & vbNewLine &_
					"There are many reason the existing code on this field in MAXIS may be incorrect and forcing this field to be manually edited every time can ensure close monitoring of this decision point."  & vbNewLine & vbNewLine &_
					"----------------------------------------"  & vbNewLine &_
					"This field will also determine:" & vbNewLine &_
					"  - Anticipated Paydates" & vbNewLine &_
					"  - Checks missing from the entry" & vbNewLine &_
					"  - Weekday of pay (for weekly and biweekly)" & vbNewLine & vbNewLine &_
					"Though this may make for slighly more work while interacting with this script, it will contribute to quality case work and actions."

'in the list checks dialog
all_checks_msg_text = 	"*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine & vbNewLine &_
						"ENTER ALL INCOME DETAIL RECEIVED/VERIFIED THROUGH ADD CHECKS BUTTON" & vbNewLine &_
						"* List all of the checks even if being excluded or the amount is $0. *" & vbNewLine &_
						"* Enter information about anticipated pay rate and hours in the area below. *" & vbNewLine &_
						"  - Typically we cannot use both but we must capture both for the notes. If both are entered, the script will guide you in chosing one to select for the prospective budget." & vbNewLine & vbNewLine &_
						"----------------------------------------" & vbNewLine &_
						"* Use the 'Add More Checks' buttons to enter checks." & vbNewLine & vbNewLine &_
						"* Use the 'Update Check Details' buttons to update the information about the check or delete it entirely." & vbNewLine & vbNewLine &_
						"* This information detail should match everything we received as verifications/documents." & vbNewLine &_
						"* The CASE:NOTE is created using all of the information here and it should include everything we know." & vbNewLine &_
						"* We have high rates of procedural and payment accuracy errors in regards to BUDGETING, CODING, and NOTING Income information. The level of accuracy and detail is handled in this script and though it may seem excessive, it meets the requirements of the programs and best serves our residents." & vbNewLine & vbNewLine &_
						"----------------------------------------" & vbNewLine &_
						"***** ADD ALL CHECKS *****" & vbNewLine &_
						" - Even if the pay amount was 0." & vbNewLine &_
						" - Even if the check will be excluded from the SNAP or CASH budget." & vbNewLine &_
						" - Even if it was determined using YTD calculations." & vbNewLine &_
						" - The script will look for checks missing in the series and connot continue without it." & vbNewLine & vbNewLine &_
						"----------------------------------------" & vbNewLine &_
						"Thank you for your attention to detail and dedication to quality and thoroughness."

'in the list checks dialog
initial_month_msg_text = 	"*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						 	"INITIAL MONTH TO UPDATE SHOULD BE THE MONTH THE CHANGE STARTED" & vbNewLine & vbNewLine &_
							"* Date of change of income is more immportant to knowing when to update than when it is verified or reported." & vbNewLine &_
							"   - For accuracy of our records, we should be going back to the month of change and updating the change informaiton." & vbNewLine &_
							"   - Updating from the month of change also allows us to use MAXIS to identify potential overpayments because the income is accurate." & vbNewLine &_
							"   - Though this may cause additional months to review and assess, this is the most complete case processing." & vbNewLine & vbNewLine &_
							"* New Jobs should be entered in the month Income began." & vbNewLine &_
							"   - This is not the month the work began, but the month the first pay was received." & vbNewLine &_
							"   - Even if the start of income does not cause an overpayment, updating the case in the month income because ensures accurate case records." & vbNewLine & vbNewLine &_
							"* Jobs that started or changed before the month of application should be updated starting in the month of application." & vbNewLine &_
							"   - If we do not have an active status in a month, we cannot update the JOBS panel." & vbNewLine & vbNewLine &_
							"UPDATE FUTURE MONTHS SHOULD BE CHECKED UNLESS THE INFORMATION IS SPECIFIC TO ONE MONTH ONLY" & vbNewLine &_
							"* Most information and changes continue for multiple months." & vbNewLine &_
							"   - It rarely makes sense to apply informaiton to only one month, changes typically continue." & vbNewLine &_
							"   - A common reason may be when updating JOBS for HRF processing as you may only have one month of RETRO income information and it is late." & vbNewLine &_
							"   - We know that there is functionality updates needed for handling END of employment and that is impact how this functionality changes." & vbNewLine & vbNewLine &_
							"* If there are two different changes that affected different months, you will need to run the script twice." & vbNewLine &_
							"   - The script cannot change the budgeting detail from month to month, it requires a new assessment and budget." & vbNewLine & vbNewLine &_
							"Remember that is coding is for each job specifically and they do not have to match each other."

pay_freq_2_msg_text		= "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						  "ABOUT ENTERING PAY FREQUENCY"  & vbNewLine & vbNewLine &_
						  "Even though you have entered the pay frequency on a previous dialog, the correct coding of this field is crucial to an accurate budget."  & vbNewLine & vbNewLine &_
						  "PAY FREQUENCY CODING ERRORS HAVE A HIGH RATE OF PAYMENT ACCURACY ERRORS."  & vbNewLine & vbNewLine &_
						  "There are many reason the existing code on this field in MAXIS may be incorrect and forcing this field to be manually edited every time can ensure close monitoring of this decision point."  & vbNewLine & vbNewLine &_
						  "----------------------------------------"  & vbNewLine &_
						  "This field will also determine:" & vbNewLine &_
						  "  - Anticipated Paydates" & vbNewLine &_
						  "  - Checks missing from the entry" & vbNewLine &_
						  "  - Weekday of pay (for weekly and biweekly)" & vbNewLine & vbNewLine &_
						  "Do not check this box without actually reviewing the information on this line."

not_thirty_msg_text		= "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						  "USING INCOME VERIFICATION THAT IS MORE OR LESS THAN 30 DAYS"  & vbNewLine & vbNewLine &_
						  "SNAP Policy used to require exactly 30 days of income verification to be sufficiently verified. This is no longer the case. While 30 days is the default, we can use any number of days from the first to last check as the correct prospective income budget. We MUST make clear why we are using something other than 30 days. This field allows for clarification of why using something OTHER than 30 days is the best budggeting decision."  & vbNewLine & vbNewLine &_
						  "Income budgeting is cause of a large portion of our errors, and many could be resolved with a clear and details CASE:NOTE on the reasoning behid the budgeting decisions applied. "  & vbNewLine & vbNewLine &_
						  "We count the 30 days by the spread of income received, not days worked." & vbNewLine &_
						  "----------------------------------------"  & vbNewLine &_
						  "Income is considered 'thirty days of income' for each pay frequency:" & vbNewLine &_
						  "  - Monthly - 30 days or fewer between first and last paychecks in the series (usually 1 check)." & vbNewLine &_
						  "  - Seemi-Monthly - Between 13 and 30 days between the first and last paychecks in the series (usually 2 checks)." & vbNewLine &_
						  "  - Biweekly - 28 days betweeen the first and last paycheecks in the series (usually 3 checks)." & vbNewLine &_
						  "  - Weekly - 28 days betweeen the first and last paycheecks in the series (usually 5 checks)." & vbNewLine & vbNewLine &_
						  "----------------------------------------"  & vbNewLine &_
						  "The field can be filled by typing in the explanation or by selecting one of the options listed in the dropdown:" & vbNewLine &_
						  "  - Income has just started and this is all that has been received." & vbNewLine &_
						  "  - Hours Reduction - this is all the income since the change." & vbNewLine &_
						  "  - Hours Increase - this is all the income since the change." & vbNewLine &_
						  "  - Wage Reduction - this is all the income since the change." & vbNewLine &_
						  "  - Wage Increase - this is all the income since the change." & vbNewLine &_
						  "  - Due to how work is scheduled, this is the best representation of expected ongoing income." & vbNewLine &_
						  "  - Client stated this income is consistent." & vbNewLine & vbNewLine &_
						  "This will require a written explanation, the more detail and clarity provided, the more likely the budget will be accepted as accurate in a review."

confirm_snap_msg_text	= "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						  "CONFIRMING THE SNAP BUDGET IS CORRECT"  & vbNewLine & vbNewLine &_
						  "This check box should ONLY be checed once you have completely reviewed the SNAP budget determined by the income information "  & vbNewLine & vbNewLine &_
						  "The budget detail listed above this checkbox, in the SNAP Budget box, was created from the information entered in the previous dialog ('All Paystubs Received') and the detail associated with each paycheck/income information."  & vbNewLine &_
						  "----------------------------------------"  & vbNewLine &_
						  "If the budget looks incorrect do NOT check the box, then press 'OK'. The script will return to the previous dialog so you can change the inputs." & vbNewLine & vbNewLine &_
						  "Some of the opptions for ways to change the budget:" & vbNewLine &_
						  "  - Ensuring all checks have been listed." & vbNewLine &_
						  "  - Checking the dates and gross amounts entered." & vbNewLine &_
						  "  - Excluding or including checks listed." & vbNewLine &_
						  "  - Changing any partial exclusions." & vbNewLine & vbNewLine &_
						  "IT IS VITAL THAT WE ARE REVIEWING THE BUDGET HERE BECUASE THE SCRIPT WILL UPDATE JOBS NEXT."

confirm_cash_msg_text	= "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						  "CONFIRMING THE CASH CHECKS ARE ACCURATE"  & vbNewLine & vbNewLine &_
						  "The checks listed are going to be entered exactly as listed in this dialog. It is important to ensure the gross amounts are listed correctly."  & vbNewLine & vbNewLine &_
						  "Since cash programs are budgeted retrospectively for earned income, the script enteres the entire gross amount in the retrospective side."  & vbNewLine & vbNewLine &_
						  "The script does not exclude paychecks in whole or partial since cash programs do not have policy to support the exclusion of checks."  & vbNewLine & vbNewLine &_
						  "----------------------------------------"  & vbNewLine &_
						  "If the checks looks incorrect do NOT check the box, then press 'OK'. The script will return to the previous dialog so you can change the inputs." & vbNewLine & vbNewLine &_
						  "IT IS VITAL THAT WE ARE REVIEWING THE BUDGET HERE BECUASE THE SCRIPT WILL UPDATE JOBS NEXT."

hc_retro_msg_text		= "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						  "HEALTH CARE RETRO BUDGET"  & vbNewLine & vbNewLine &_
						  "The script has the ability to update JOBS with the procedure for prospective or retrospective budgeting. The default for health care budgets is to use a prospecitve budget using the pay amount listed in the HC Income Estimate Pop-Up."  & vbNewLine & vbNewLine &_
						  "The only way to force the script to update the retrospective side of the JOBS panel is to check this box"  & vbNewLine & vbNewLine &_
						  "This is typically used for Manual Monthly Spenddown cases - often LTC cases. If you are unsure if this should apply, contact Knowledge Now."

confirm_hc_msg_text		= "*** Tips and Tricks ***" & vbNewLine & "----------------------------------------" & vbNewLine &_
						  "CONFIRMING THE HC INCOME ESTIMATION AND BUDGET IS CORRECT"  & vbNewLine & vbNewLine &_
						  "The health care budget is determined by the check dates and the calculated amount of the HC Income Estimate."  & vbNewLine & vbNewLine &_
						  "The HC Income Estimate is based on the average of the provided checks."  & vbNewLine & vbNewLine &_
						  "----------------------------------------"  & vbNewLine &_
						  "If the budget looks incorrect do NOT check the box, then press 'OK'. The script will return to the previous dialog so you can change the inputs." & vbNewLine & vbNewLine &_
						  "Some of the opptions for ways to change the budget:" & vbNewLine &_
						  "  - Ensuring all checks have been listed." & vbNewLine &_
						  "  - Checking the dates and gross amounts entered." & vbNewLine & vbNewLine &_
						  "IT IS VITAL THAT WE ARE REVIEWING THE BUDGET HERE BECUASE THE SCRIPT WILL UPDATE JOBS NEXT."


'THE SCRIPT ================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call check_for_MAXIS(False)
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

original_month = MAXIS_footer_month     'setting these to a seperate variable for the dialog
original_year = MAXIS_footer_year

future_months_check = checked           'default to having th script update future months

'INITIAL Dialog - case number, footer month, worker signature
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 190, 245, "ACTIONS - Earned Income Budgeting"
	EditBox 90, 5, 70, 15, MAXIS_case_number
	EditBox 100, 25, 15, 15, original_month
	EditBox 120, 25, 15, 15, original_year
	CheckBox 10, 45, 140, 10, "Check here to have the script update all", future_months_check
	EditBox 5, 80, 175, 15, worker_signature
	ButtonGroup ButtonPressed
		PushButton 140, 25, 15, 15, "!", tips_and_tricks_button
		PushButton 15, 202, 85, 13, "FULL INSTRUCTIONS", instructions_btn
		PushButton 100, 202, 25, 13, "FAQ", faq_btn
		PushButton 125, 202, 50, 13, "Quick Start", quick_start_btn
		OkButton 80, 225, 50, 15
		CancelButton 135, 225, 50, 15
	Text 5, 10, 85, 10, "Enter the case number:"
	Text 5, 30, 90, 10, "Starting Footer Month/Year:"
	Text 20, 55, 120, 10, "future months and send through BG."
	Text 5, 70, 65, 10, "Worker Signature:"
	GroupBox 5, 100, 180, 120, "INSTRUCTIONS - PLEASE READ!!!"
	Text 10, 115, 170, 20, "This script is used to enter EARNED Income onto JOBS panels when verified and ready for budgeting."
    Text 10, 140, 170, 20, "FUNCTIONALITY: create a panel, used actual check, use a verified anticipated income amount."
    Text 10, 165, 170, 35, "PRIMARY PURPOSE: Accurate detailing of budgeting descions, including exclusion of checks, partial checks, selection between actual or anticipated budgets."
	' Text 10, 150, 170, 40, "If a JOBS panel or BUSI panel needs to be added to MAXIS for a client or income source, the script will ask for any panels that need to be added first. Review the case now to ensure that the correct action will be taken in the correct order."
EndDialog

'calling the dialog
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation

        If IsNumeric(MAXIS_case_number) = FALSE or Len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "* Enter a valid case number."       'confirming a valid case number
        If trim(worker_signature) = "" Then err_msg = err_msg & vbNewLine & "* Enter your worker signature for your case notes."                        'confirming there is a worker signature

        original_month = trim(original_month)       'cleaning up the entry here
        original_year = trim(original_year)
        If len(original_year) <> 2 or len(original_month) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a 2 digit footer month and year."          'forcing 2 digit month and year to be entered

		If ButtonPressed = tips_and_tricks_button Then
			tips_tricks_msg = MsgBox("*** Tips and Tricks ***" & vbNewLine & "--------------------" & vbNewLine & vbNewLine &_
									"Selecting the starting footer month and year will determine what month the script will search for any existing JOBS panels. Select the month that has the JOBS panels you need to take action on." & vbNewLine & vbNewLine &_
									"The Footer Month is able to be changed for each job specifically later in the script as every job is not going to have the same months to change." & vbNewLine & vbNewLine &_
									"Remember to update JOBS in the month of change not just the current month or next month."  & vbNewLine & vbNewLine &_
									"The script may select the footer month based on some of the information provided such as income start date or pay dates.", vbInformation, "Tips and Tricks")
									' ""  & vbNewLine & vbNewLine &_
			err_msg = "LOOP" & err_msg
		End If
		If ButtonPressed = instructions_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20EARNED%20INCOME%20BUDGETING.docx"
		If ButtonPressed = faq_btn          Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20EARNED%20INCOME%20BUDGETING%20-%20FAQ.docx"
		If ButtonPressed = quick_start_btn  Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20EARNED%20INCOME%20BUDGETING%20QUICK%20START.docx"
		If ButtonPressed = instructions_btn OR ButtonPressed = faq_btn OR ButtonPressed = quick_start_btn Then err_msg = "LOOP" & err_msg

        If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "-- Please resolve the following to continue --" & vbNewLine & err_msg                                             'displaying the error handling
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

MAXIS_footer_month = original_month     'setting the footer month and year back to what was entered in the dialog.
MAXIS_footer_year = original_year       'this is split out for the option of having seperate handling prior to the reassignment for working in current month if needed

Call back_to_SELF                       'need to gather some detail to have the correct script run

developer_mode = FALSE                  'allowing worker to exit if started in Inquiry on accident
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
    developer_mode = TRUE
End If
If developer_mode = TRUE then MsgBox "Developer Mode ACTIVATED!"        'developer mode difference is that the MAXIS update detail is shown in a messagebox instead of updating the panel
If developer_mode = TRUE Then script_run_lowdown = "Run in INQUIRY"

'If we are not running in developer mode - making sure we can edit STAT panels on this case
If developer_mode = False Then
	Call navigate_to_MAXIS_screen("STAT", "ADDR")		'want to use a panel that every case has and there is only one of
	PF9													'Try to put it in EDIT Mode
	EMReadScreen edit_mode_code, 1, 20, 8
	If edit_mode_code <> "E" Then
		EMReadScreen MAXIS_edit_mode_error_message, 78, 24, 2
		MAXIS_edit_mode_error_message = trim(MAXIS_edit_mode_error_message)
		cannot_edit_end_msg = "It appears this case cannot be updated in STAT at this time." & vbCr & vbCr &_
							"* * * * MAXIS message * * * *" & vbCr & vbCr & MAXIS_edit_mode_error_message & vbCr & vbCr &_
							"-----------------------------------------------------------------------------" & vbCr & vbCr &_
							"The EARNED INCOME BUDGETING Script will now end." & vbCr &_
							"Review the case and try again once STAT panels can be updated."
		script_end_procedure_with_error_report(cannot_edit_end_msg)
	Else
		PF10 											'Oops out of the edit
		Call back_to_SELF
	End If
End If

vars_filled = FALSE
Call restore_your_work(vars_filled)			'looking for a 'restart' run

If vars_filled = False Then

	Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)                           'Going to find the FS Application Date
	If is_this_priv = True Then call script_end_procedure("This script will now end because it appears this case is privileged.")
	curr_row = 1
	curr_col = 1
	EMSearch " FS:", curr_row, curr_col
	If curr_row <> 0 Then
		EMReadScreen fs_prog_status, 7, curr_row, curr_col + 5
		fs_prog_status = trim(fs_prog_status)
		If fs_prog_status <> "" Then
			EMReadScreen fs_appl_date, 8, curr_row, curr_col + 25
            If NOT IsDate(fs_appl_date) Then fs_appl_date = ""
		End If
	End If
	fs_appl_footer_month = left(fs_appl_date, 2)                            'Making a footer month and year for the FS Application'
	fs_appl_footer_year = right(fs_appl_date, 2)

	Do                                                                      'Getting in to STA (making sure we get past background)
		Call navigate_to_MAXIS_screen ("STAT", "SUMM")
		EMReadScreen summ_check, 4, 2, 46
	Loop until summ_check = "SUMM"

	CASH_case = FALSE       'defining these as a default
	SNAP_case = FALSE
	HC_case = FALSE

	Call Navigate_to_MAXIS_screen("STAT", "PROG")                           'Getting program status to identify potential programs the income should apply to

	EMReadScreen cash_one_status, 4, 6, 74                                  'reading each program status
	EMReadScreen cash_two_status, 4, 7, 74
	EMReadScreen cash_one_prog, 2, 6, 64                                  'reading each program status
	EMReadScreen cash_two_prog, 2, 7, 64
	EMReadScreen grh_status, 4, 9, 74
	EMReadScreen snap_status, 4, 10, 74
	EMReadScreen hc_status, 4, 12, 74

	If cash_one_status = "ACTV" OR cash_one_status = "PEND" Then CASH_case = TRUE   'setting programs to TRUE based on PROG status
	If cash_two_status = "ACTV" OR cash_two_status = "PEND" Then CASH_case = TRUE
	If grh_status = "ACTV" OR grh_status = "PEND" Then GRH_case = TRUE
	If snap_status = "ACTV" OR snap_status = "PEND" Then SNAP_case = TRUE
	If hc_status = "ACTV" OR hc_status = "PEND" Then HC_case = TRUE


	panels_count = 0								'this is our counter to add new panels to the EARNED_INCOME_PANELS_ARRAY
	panel_exists = False

	' call HH_member_custom_dialog(HH_member_array)   'finding who should be looked at for income on the case
    'Since the panels are optional later - we do not need the HH member dialog
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
    transmit

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
        EMReadScreen access_denied_check, 13, 24, 2
        If access_denied_check = "ACCESS DENIED" Then
            PF10
			EMWaitReady 0, 0
        Else
    		EMReadscreen ref_nbr, 3, 4, 33
            client_string = trim(ref_nbr)
            client_array = client_array & client_string & " "
        End If
		transmit
	    Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	client_array = TRIM(client_array)
	test_array = split(client_array)
	ReDim HH_member_array(Ubound(test_array))			'setting the upper bound for how many spaces to use from the array
    For duck = 0 to UBound(test_array)
        HH_member_array(duck) = test_array(duck)       'filling the array with the strings from the client_array
    Next


	'FUTURE FUNCTIONALITY - Stop work should be added in before we add information to the EARNED_INCOME_PANELS_ARRAY

	Call navigate_to_MAXIS_screen("STAT", "JOBS")       'Starting with JOBS panels
	For each member in HH_member_array                  'We are going to look at each HH member checked in the HH_member dialog
		EMWriteScreen member, 20, 76                    'going to the member in JOBS
		Transmit

		EMReadScreen number_of_jobs_panels, 1, 2, 78    'finding the total number of panels currently existing for the current member.

		If number_of_jobs_panels <> "0" Then            'if there are 0 panels we don't need to do anything else in JOBS for this member
			number_of_jobs_panels = number_of_jobs_panels * 1       'making the number read and actual number

			For panel = 1 to number_of_jobs_panels      'we are going to cycle through each of the panels for this member
				EMWriteScreen "0" & panel, 20, 79       'navigating to the panel instance
				transmit
				'FUTURE FUNCTIONALITY - Stop work should be added in before we add information to the EARNED_INCOME_PANELS_ARRAY

				save_this_panel = TRUE                  'we are always at this point going to save the panel to the EARNED_INCOME_PANELS_ARRAY
														'FUTURE FUNCTIONALITY where we may be deleting old panels, in which case we would NOT be saving the panel to the array

				EMReadScreen end_date, 8, 9, 49         'finding the end date

				If end_date <> "__ __ __" Then
					end_date = replace(end_date, " ", "/")
					end_date = DateValue(end_date)

					'FUTURE FUNCTIONALITY - here is a start with deleting an already ended panel - this has not been tested or completed
					'TESTING NEEDED
					' If DateDiff("m", end_date, date) > 3 Then
					'
					'     Dialog1 = ""
					'     BeginDialog Dialog1, 0, 0, 186, 140, "Dialog"
					'       OptionGroup RadioGroup1
					'         RadioButton 20, 45, 70, 10, "Delete this Panel", delete_panel
					'         RadioButton 20, 60, 75, 10, "Leave this Panel", leave_ended_panel
					'       EditBox 10, 100, 170, 15, explain_leaving_ended_panel
					'       ButtonGroup ButtonPressed
					'         OkButton 130, 120, 50, 15
					'       Text 10, 10, 170, 25, "This JOBS panel indicates the income ended more than 3 months ago. This panel is no longer needed in this month since this income has ended."
					'       Text 10, 80, 115, 20, "If this ened panel is to be left, explain why it is still needed:"
					'     EndDialog
					'
					'     Do
					'         Do
					'             err_msg = ""
					'
					'             Dialog Dialog1
					'
					'             explain_leaving_ended_panel = trim(explain_leaving_ended_panel)
					'
					'             If leave_ended_panel = checked and explain_leaving_ended_panel = "" Then err_msg = err_msg & vbNewLine & "* If an ended panel is to be left on an active case, explain why it is still needed."
					'
					'             If err_msg <> "" Then MsgBox "** Please Resolve to Continue **" & vbNewLine & err_msg
					'         Loop Until err_msg = ""
					'         call check_for_password(are_we_passworded_out)
					'     Loop until are_we_passworded_out = false
					'
					'     If delete_panel = checked then panels_to_delete = panels_to_delete & "~" & "JOBS " & member & " " & "0" & panel
					'
					' End If
				End If

				If save_this_panel = TRUE Then                                      'if the panel will be saved (always for now) then we are going to read panel detail.
                    panel_exists = True
					ReDim Preserve JOBS_PANELS(panels_count)									'resizing the array
					Set JOBS_PANELS(panels_count) = new jobs_income

					'Setting known information and defaults
					JOBS_PANELS(panels_count).member = member                'member known from member array
					JOBS_PANELS(panels_count).instance = "0" & panel         'instance known from the for-next of all panels for this member
					JOBS_PANELS(panels_count).income_received = False
					If CASH_case = TRUE Then JOBS_PANELS(panels_count).apply_to_CASH = checked     'These are defaulted by whatever program is active or pending - will be able to be changed later
					If SNAP_case = TRUE Then JOBS_PANELS(panels_count).apply_to_SNAP = checked
					If HC_case = TRUE Then JOBS_PANELS(panels_count).apply_to_HC = checked
					If GRH_case = TRUE Then JOBS_PANELS(panels_count).apply_to_GRH = checked

					Call JOBS_PANELS(panels_count).read_panel				'capture current panel information

					panels_count = panels_count + 1       'incrementing our counter to be ready for the next panel/member/income type
				End If      'If save_this_panel = TRUE Then
			Next            'For panel = 1 to number_of_jobs_panels
		End If              'If number_of_jobs_panels <> "0" Then
	Next                    'For each member in HH_member_array
	' panel_exists = TRUE     'default to panels existing - this is also reset on each loop so that if a panel is added this statement is reassessed.

	Call offer_new_panel_creation

	If panel_exists = FALSE Then script_end_procedure("There are no earned income panels on this case to update. Run the script again and add a panel or review the case and verifications you are updating.")
	' save_your_work
End If

								'----------------------------------------------------------'
                        '---------------------------------------------------------------------------------'
'------------------------------------------------- GATHERING PAY INFORMATION FOR EACH PANEL --------------------------------------------------'
                        '---------------------------------------------------------------------------------'
                                '----------------------------------------------------------'

Call back_to_SELF               'this is a good reset
Do                              'getting back in to STAT in case it went to background adding a new panel.
	Call navigate_to_MAXIS_screen ("STAT", "SUMM")
	EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"

pay_item = 0                    'counter for adding items to the LIST_OF_INCOME_ARRAY
For panel = 0 to UBOUND(JOBS_PANELS)       'looping through all of the current JOBS or BUSI panels
    cancel_clarify = ""         'blanking this from previous use AND from loops through the panels

	Call Navigate_to_MAXIS_screen("STAT", "JOBS")                       'navigate to the current panel in the array
	EMWriteScreen JOBS_PANELS(panel).member, 20, 76
	EMWriteScreen JOBS_PANELS(panel).instance, 20, 79
	transmit

	If vars_filled = False Then
		'The script can only update JOBS panels if there is no income end date.
		If JOBS_PANELS(panel).income_end_dt = "" Then
			'This is where the worker indicates if they have income information to budget for this panel.
			'If they click 'No' here, there is no way to go back to this question for that panel.
			employer_check = MsgBox("Do you have income verification for this job? Employer name: " & JOBS_PANELS(panel).employer, vbYesNo + vbQuestion, "Select Income Panel")
			If employer_check = vbYes Then script_run_lowdown = script_run_lowdown & vbCr & "Income received for - " & JOBS_PANELS(panel).employer & " button pressed - YES"
			If employer_check = vbNo Then script_run_lowdown = script_run_lowdown & vbCr & "Income received for - " & JOBS_PANELS(panel).employer & " button pressed - NO"
		Else
			'If there is already an end date the script will need to force that it cannot be updated.
			MsgBox "This job appears to be ended." & vbNewLine & vbNewLine & "The employer name: " & JOBS_PANELS(panel).employer & vbNewLine & "End Date: " & JOBS_PANELS(panel).income_end_dt & vbNewLine & vbNewLine & "This script connot accomodate the update of this panel at this time. This should be processed manually." & vbNewLine & "If this job has not ended, once this script run is completed, remove the income end date and rerun the script for this job."
			script_run_lowdown = script_run_lowdown & vbCr & "Employer - " & JOBS_PANELS(panel).employer & " has an income end date of " & JOBS_PANELS(panel).income_end_dt & " and cannot be updated by the script at this time."
			employer_check = vbNo
		End If
		'Some panels will have this defaulted  already but if not, this will defalt to the footer month and year inidicated in the initial dialog
		If JOBS_PANELS(panel).initial_month_mo = "" Then JOBS_PANELS(panel).initial_month_mo = MAXIS_footer_month
		If JOBS_PANELS(panel).initial_month_yr = "" Then JOBS_PANELS(panel).initial_month_yr = MAXIS_footer_year
		JOBS_PANELS(panel).update_future = future_months_check      'defaulted to the checkbox in the initial dialog
		JOBS_PANELS(panel).EI_panel_vbYes = False
		If employer_check = vbYes Then JOBS_PANELS(panel).EI_panel_vbYes = True
	End If

	'If the worker indicates there is income information for this panel.
	If JOBS_PANELS(panel).EI_panel_vbYes = True Then
        If vars_filled = False Then
            'There are a bunch of variables and information to default.
            JOBS_PANELS(panel).income_received = TRUE        'This is set to indicate actions are needed by the rest of the script
            review_small_dlg = TRUE                                             'The ENTER PAY INFORMATION Dialog is only shown if this is 'TRUE'
            JOBS_PANELS(panel).ignore_antic = FALSE          'This will be determined TURE as needed later if appropriate

            JOBS_PANELS(panel).add_check
        End If
        JOBS_PANELS(panel).budget_confirmed = False
		Do
			Do
				If JOBS_PANELS(panel).budget_confirmed = False Then JOBS_PANELS(panel).update_job_detail
                save_your_work()

				JOBS_PANELS(panel).confirm_job_budget()
                save_your_work()
			Loop until JOBS_PANELS(panel).budget_confirmed = True
			Call check_for_password(are_we_passworded_out)  'we are doing password handling before error handling because of the 2 dialogs looped together
		Loop until are_we_passworded_out = False

	End If
    save_your_work()
Next

'WE PREVIOUSLY HAD FUNCTIONALITY TO SEE IF THE SNAP WAS A UHFS CASE FOR RETRO BUDGETING - BUT THIS SHOULD NOT BE NEEDED AS OF 3/1/2025

                                '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
'------------------------------------------------- DETERMINING WHICH MONTHS TO UPDATE --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                                '----------------------------------------------------------'

list_of_all_months_to_update = "~"      'start of a list that will become and array
update_with_verifs = FALSE              'defaults to false

For panel = 0 to UBOUND(JOBS_PANELS)       'looping through all of the current JOBS or BUSI panels

    If JOBS_PANELS(panel).income_received = TRUE Then       'only looking at panels that have income received
        update_with_verifs = TRUE                           'if any panel has income received, we will be updating with verifications
        JOBS_PANELS(panel).update_this_month = False        'defaulting this for each panel - this will be updated for each month on each panel when the updating actually happens

        'here we look at the initial month for each panel - making a date for the first of the initial month
        mm_1_yy = JOBS_PANELS(panel).initial_month_mo & "/1/" & JOBS_PANELS(panel).initial_month_yr
        mm_1_yy = DateValue(mm_1_yy)
        If InStr(list_of_all_months_to_update, "~" & mm_1_yy & "~") = 0 Then    'looks to see if the initial month has already been added to the list on a previous loop
            list_of_all_months_to_update = list_of_all_months_to_update & mm_1_yy & "~" 'if not, it is added to the list
        End If

        If JOBS_PANELS(panel).update_future = checked Then      'if the panel is set to update future months
            next_month = mm_1_yy        'this is the initial month to start with
            CM_plus_2 = DateValue(CM_plus_2_mo & "/1/" & CM_plus_2_yr)          'setting a date for current month plus 2
            CM_plus_2 = DateValue(CM_plus_2)
            Do          'now we are going to loop to keep adding months
                If InStr(list_of_all_months_to_update, "~" & next_month & "~") = 0 Then     'if the month is NOT in the list, add it
                    list_of_all_months_to_update = list_of_all_months_to_update & next_month & "~"
                End If

                next_month = DateAdd("m", 1, next_month)                        'now increment to the next month
            Loop until next_month = CM_plus_2                                   'stop one the next month is current month plus 2 because AMXIS doesn't go there
        End If
    End If

    Call back_to_SELF
Next

                            '----------------------------------------------------------'
                    '---------------------------------------------------------------------------------'
    '-------------------------------------------------GOING TO UPDATE THE PANEL --------------------------------------------------'
                    '---------------------------------------------------------------------------------'
                            '----------------------------------------------------------'

If update_with_verifs = TRUE Then       'this means we have at least one panel with income to update'
    list_of_all_months_to_update = right(list_of_all_months_to_update, len(list_of_all_months_to_update)-1)     'formatting our list of months
    list_of_all_months_to_update = left(list_of_all_months_to_update, len(list_of_all_months_to_update)-1)
    If InStr(list_of_all_months_to_update, "~") <> 0 Then                       'making the list an array in chronologival order
        update_months_array = split(list_of_all_months_to_update, "~")
        Call sort_dates(update_months_array)
    Else
        update_months_array = array(list_of_all_months_to_update)
    End If
    script_run_lowdown = script_run_lowdown & vbCr & "The months to update - " & list_of_all_months_to_update & vbCr & vbCr & "UPDATE DETAILS:" & vbCr

    Call back_to_SELF       'reset

    next_cash_month = 0         'setting this because we have ANOTHER ARRAY!
    For each active_month in update_months_array                'now we loop through the list of months we found before
        script_run_lowdown = script_run_lowdown & "Footer Month: " & active_month & " ----------"

        For panel = 0 to UBOUND(JOBS_PANELS)       'looping through all of the current JOBS or BUSI panels
            JOBS_PANELS(panel).updates_to_display = ""

            EMReadScreen summ_check, 4, 2, 46                       'Making sure we start at SUMM
            'BUGGY CODE - need to make sure we are also in the right footer month here
            If summ_check <> "SUMM" Then                            'at the end of the loop we go to summ so we should be already there
                Call back_to_SELF
                Do
                    Call navigate_to_MAXIS_screen("STAT", "SUMM")
                    EMReadScreen summ_check, 4, 2, 46
                Loop until summ_check = "SUMM"
            End If


            If JOBS_PANELS(panel).income_received = TRUE Then       'only looking at panels that have income received
                'BUGGY CODE - this may miss the first check(s) in a month if they are not listed or the pay date is after the first one for the first month

                JOBS_PANELS(panel).update_panel(active_month)                                                   'ALL THE JUICY BITS GO HERE

                If JOBS_PANELS(panel).updates_to_display <> "" AND developer_mode = TRUE Then MsgBox JOBS_PANELS(panel).updates_to_display            'this shows the information that WOULD have been updated if we were not in INQUIRY
				script_run_lowdown = script_run_lowdown & vbCr & JOBS_PANELS(panel).updates_to_display

                'If this panel is should to update months after the initial month, this is saved for the next loop to have it updated
                'FUTURE FUNCTIONALITY - if we need to change how we handle the future month updates thing or dealing with STWK - this would be here
                If JOBS_PANELS(panel).update_future = unchecked Then JOBS_PANELS(panel).update_this_month = False

			End If
            EMWriteScreen "SUMM", 20, 71        'go back to SUMM'
            transmit
            EMReadScreen no_hours, 40, 6, 16
            If no_hours = "PROSPECTIVE EARNINGS EXIST WITH NO HOURS" Then
                EMWriteScreen "Y", 9, 58
                transmit
            End If
        Next            'For ei_panel = 0 to UBOUND(EARNED_INCOME_PANELS_ARRAY, 2)

		transmit        'after all of the panels have been reviewed we are going to STAT/WRAP to get to the next month without sending through background if possible

        EmWriteScreen "BGTX", 20, 71
        transmit
        If active_month <> update_months_array(ubound(update_months_array)) Then EmWriteScreen "Y", 16, 54      'if we are at the last month to update, then we leave
        transmit
        'BUGGY CODE - need to add a check here that we made it to STAT/WRAP and that we are in the next footer month

        If active_month <> update_months_array(ubound(update_months_array)) Then        'getting to SUMM
            EmWriteScreen "SUMM", 20, 71
            transmit
        End If


    Next            'For each active_month in update_months_array
End If          'If update_with_verifs = TRUE Then

                '----------------------------------------------------------'
        '---------------------------------------------------------------------------------'
'-------------------------------------------------CASE NOTING --------------------------------------------------'
        '---------------------------------------------------------------------------------'
                '----------------------------------------------------------'
end_msg = ""

For panel = 0 to UBOUND(JOBS_PANELS)       'looping through all of the current JOBS or BUSI panels
    JOBS_PANELS(panel).updates_to_display = ""
    JOBS_PANELS(panel).set_programs()

    If JOBS_PANELS(panel).new_panel = TRUE OR JOBS_PANELS(panel).income_received = TRUE Then
        'updating information for when the script ends
        end_msg = end_msg & vbNewLine & "Updated JOBS for MEMB " & JOBS_PANELS(panel).member & " at " & JOBS_PANELS(panel).employer
        If JOBS_PANELS(panel).new_panel = TRUE Then end_msg = end_msg & " panel added eff with start date " & JOBS_PANELS(panel).income_start_dt
        If JOBS_PANELS(panel).income_received = TRUE Then end_msg = end_msg & " income budgeted, panel updated."

        JOBS_PANELS(panel).case_note_details(developer_mode)
    End If
Next

If end_msg = "" Then end_msg = "Script ended with no action taken, panels not updated, no case note created. No new panels were indicated and no income verification was entered to be budgeted."

script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/19/2025
'--Tab orders reviewed & confirmed----------------------------------------------09/19/2025
'--Mandatory fields all present & Reviewed--------------------------------------09/19/2025
'--All variables in dialog match mandatory fields-------------------------------09/19/2025
'Review dialog names for content and content fit in dialog----------------------09/19/2025
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------09/19/2025
'--Create a button to reference instructions------------------------------------09/19/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/19/2025
'--CASE:NOTE Header doesn't look funky------------------------------------------09/19/2025
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/19/2025
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------09/19/2025
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/19/2025
'--MAXIS_background_check reviewed (if applicable)------------------------------09/19/2025
'--PRIV Case handling reviewed -------------------------------------------------09/19/2025
'--Out-of-County handling reviewed----------------------------------------------09/19/2025
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/19/2025
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------09/19/2025
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/19/2025
'--Incrementors reviewed (if necessary)-----------------------------------------09/19/2025
'--Denomination reviewed -------------------------------------------------------09/19/2025
'--Script name reviewed---------------------------------------------------------09/19/2025
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/06/2025
'--comment Code-----------------------------------------------------------------09/19/2025
'--Update Changelog for release/update------------------------------------------09/19/2025
'--Remove testing message boxes-------------------------------------------------09/19/2025
'--Remove testing code/unnecessary code-----------------------------------------09/19/2025
'--Review/update SharePoint instructions----------------------------------------10/06/2025
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/06/2025
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------10/06/2025
'--Update project team/issue contact (if applicable)----------------------------10/06/2025
