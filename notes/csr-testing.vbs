'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - TESTING - CSR.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 600          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================
' run_locally = TRUE
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
call changelog_update("09/30/2020", "Changed the 'Details Blank' checkbox to being embeded in the droplist.", "Casey Love, Hennepin County")
call changelog_update("09/09/2020", "Added a new dialog between the form completion and the MAXIS Detail and Information entry to make the process clearer.", "Casey Love, Hennepin County")
call changelog_update("08/27/2020", "Removed the first dialog of questions and incorporated them into the detail entry dialogs.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
call changelog_update("12/22/2018", "Added closing message reminder about accepting all ECF work items for CSR's at the time of processing.", "Ilse Ferris, Hennepin County")
call changelog_update("12/07/2018", "Added Paperless (*) IR Option back, with updated functionality.", "Casey Love, Hennepin County")
call changelog_update("11/27/2018", "Removed Paperless (*) IR Option as this CASE/NOTE was insufficient.", "Casey Love, Hennepin County")
call changelog_update("01/17/2017", "This script has been updated to clean up the case note. The script was case noting the ''Verifs Needed'' section twice. This has been resolved.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function access_JOBS_panel(access_type, job_member, job_verif, job_employer, job_type, job_pay_amount, job_prosp_total, job_prosp_hours, job_frequency, job_update_date, job_start_date, job_end_date, panel_ref_numb, hourly_wage, retrospective_total, retrospective_hours, fs_pic_pay_frequency, fs_pic_average_hours, fs_pic_average_pay, fs_pic_monthly_prospective, grh_pic_pay_frequency, grh_pic_average_pay, grh_pic_monthly_prospective, jobs_subsidy_code)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen job_member, 2, 4, 33
        EMReadScreen job_type, 1, 5, 34
        EMReadScreen jobs_subsidy_code, 2, 5, 74
        EMReadScreen job_verif, 1, 6, 34
        EMReadScreen hourly_wage, 6, 6, 75
        EMReadScreen employer_name, 30, 7, 42
        EMReadScreen income_start_date, 8, 9, 35
        EMReadScreen income_end_date, 8, 9, 49
        EMReadScreen retrospective_total, 8, 17, 38
        EMReadScreen retrospective_hours, 3, 18, 43

        For jobs_row = 16 to 12 Step -1
            EMReadScreen paycheck_amount, 8, jobs_row, 67
            If paycheck_amount <> "________" Then
                job_pay_amount = trim(paycheck_amount)
                Exit For
            End If
        Next
        EMReadScreen job_prosp_total, 8, 17, 67
        EMReadScreen job_frequency, 1, 18, 35
        EMReadScreen job_prosp_hours, 3, 18, 72
        EMReadScreen last_updated, 8, 21, 55
        ' MsgBox "Line 817" & vbNewLine & last_updated

        EMWriteScreen "X", 19, 38           'opening the FS PIC
        transmit

        EMReadScreen fs_pic_pay_frequency, 1, 5, 64
        EMReadScreen fs_pic_average_hours, 6, 16, 51
        EMReadScreen fs_pic_average_pay, 8, 17, 56
        EMReadScreen fs_pic_monthly_prospective, 8, 18, 56

        PF3                                 'closing the FS PIC

        EMWriteScreen "X", 19, 71           'opening the GRH PIC
        transmit

        EMReadScreen grh_pic_pay_frequency, 1, 3, 63
        EMReadScreen grh_pic_average_pay, 8, 16, 65
        EMReadScreen grh_pic_monthly_prospective, 8, 17, 65

        PF3                                 'closing the GRH PIC

        If jobs_subsidy_code = "01" Then jobs_subsidy_code = "01 - Subsidized Public Secotr Employer"
        If jobs_subsidy_code = "02" Then jobs_subsidy_code = "02 - Subsidized Private Sector Employer"
        If jobs_subsidy_code = "03" Then jobs_subsidy_code = "03 - On-the-Job-Training"
        If jobs_subsidy_code = "04" Then jobs_subsidy_code = "04 - Americorps"
        If jobs_subsidy_code = "__" Then jobs_subsidy_code = "None"

        hourly_wage = trim(hourly_wage)
        retrospective_total = trim(retrospective_total)
        retrospective_hours = trim(retrospective_hours)

        job_employer = replace(employer_name, "_", "")
        If job_verif = "1" Then job_verif = "1 - Pay Stubs"
        If job_verif = "2" Then job_verif = "2 - Empl Stmt"
        If job_verif = "3" Then job_verif = "3 - Coltrl Stmt"
        If job_verif = "4" Then job_verif = "4 - Other Doc"
        If job_verif = "5" Then job_verif = "5 - Pend Out State"
        If job_verif = "N" Then job_verif = "N - No Verif Prvd"
        If job_verif = "?" Then job_verif = "? - Delayed Verif"

        If job_type = "J" Then job_type = "J - WIOA"
        If job_type = "W" Then job_type = "W - Wages"
        If job_type = "E" Then job_type = "E - EITC"
        If job_type = "G" Then job_type = "G - Experience Works"
        If job_type = "F" Then job_type = "F - Fed Work Study"
        If job_type = "S" Then job_type = "S - State Work Study"
        If job_type = "O" Then job_type = "O - Other"
        If job_type = "C" Then job_type = "C - Contract Income"
        If job_type = "T" Then job_type = "T - Training Prog"
        If job_type = "P" Then job_type = "P - Service Prog"
        If job_type = "R" Then job_type = "R - Rehab Prog"

        job_prosp_total = trim(job_prosp_total)
        job_prosp_hours = trim(job_prosp_hours)
        If job_frequency = "1" Then job_frequency = "1 - Monthly"
        If job_frequency = "2" Then job_frequency = "2 - Semi Monthly"
        If job_frequency = "3" Then job_frequency = "3 - Biweekly"
        If job_frequency = "4" Then job_frequency = "4 -  Weekly"
        If job_frequency = "5" Then job_frequency = "5 - Other"

        job_update_date = replace(last_updated, " ", "/")
        ' MsgBox "Line 849" & vbNewLine & job_update_date
        job_start_date = replace(income_start_date, " ", "/")
        job_end_date = replace(income_end_date, " ", "/")
        if job_end_date = "__/__/__" then job_end_date = ""

        If fs_pic_pay_frequency = "1" Then fs_pic_pay_frequency = "1 - Monthly"
        If fs_pic_pay_frequency = "2" Then fs_pic_pay_frequency = "2 - Semi Monthly"
        If fs_pic_pay_frequency = "3" Then fs_pic_pay_frequency = "3 - Biweekly"
        If fs_pic_pay_frequency = "4" Then fs_pic_pay_frequency = "4 -  Weekly"
        If fs_pic_pay_frequency = "5" Then fs_pic_pay_frequency = "5 - Other"
        If fs_pic_pay_frequency = "_" Then fs_pic_pay_frequency = ""

        fs_pic_average_hours = trim(fs_pic_average_hours)
        fs_pic_average_pay = trim(fs_pic_average_pay)
        fs_pic_monthly_prospective = trim(fs_pic_monthly_prospective)

        If grh_pic_pay_frequency = "1" Then grh_pic_pay_frequency = "1 - Monthly"
        If grh_pic_pay_frequency = "2" Then grh_pic_pay_frequency = "2 - Semi Monthly"
        If grh_pic_pay_frequency = "3" Then grh_pic_pay_frequency = "3 - Biweekly"
        If grh_pic_pay_frequency = "4" Then grh_pic_pay_frequency = "4 -  Weekly"
        If grh_pic_pay_frequency = "5" Then grh_pic_pay_frequency = "5 - Other"
        If grh_pic_pay_frequency = "_" Then grh_pic_pay_frequency = ""

        grh_pic_average_pay = trim(grh_pic_average_pay)
        grh_pic_monthly_prospective = trim(grh_pic_monthly_prospective)

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_BUSI_panel(access_type, busi_member, busi_type, income_start_date, income_end_date, cash_net_prosp_amount, cash_net_retro_amount, cash_retro_total_income, cash_retro_expenses, cash_prosp_total_income, cash_prosp_expenses, cash_income_verif, cash_expense_verif, snap_net_prosp_amount, snap_net_retro_amount, snap_retro_total_income, snap_retro_expenses, snap_prosp_total_income, snap_prosp_expenses, snap_income_verif, snap_expense_verif, hc_method_a_net_prosp_amount, hc_method_b_net_prosp_amount, hc_method_a_total_income, hc_method_a_expenses, hc_method_a_income_verif, hc_method_a_expense_verif, hc_method_b_total_income, hc_method_b_expenses, hc_method_b_income_verif, hc_method_b_expense_verif, SE_method, SE_method_date, reported_hours, minimum_wage_hours, update_date, panel_ref_numb)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen busi_member, 2, 4, 33
        EMReadScreen type_of_income, 2, 5, 37
        EMReadScreen income_start_date, 8, 5, 55
        EMReadScreen income_end_date, 8, 5, 72
        EMReadScreen cash_net_prosp_amount, 8, 8, 69
        EMReadScreen cash_net_retro_amount, 8, 8, 55
        EMReadScreen snap_net_prosp_amount, 8, 10, 69
        EMReadScreen snap_net_retro_amount, 8, 10, 55
        EMReadScreen hc_method_a_net_prosp_amount, 8, 11, 69
        EMReadScreen hc_method_b_net_prosp_amount, 8, 12, 69
        EMReadScreen reported_hours, 3, 13, 74
        EMReadScreen minimum_wage_hours, 3, 14, 74
        EMReadScreen update_date, 8, 21, 55

        EMReadScreen SE_method, 2, 16, 53
        EMReadScreen SE_method_date, 8, 16, 63

        ' MsgBox "Line 870" & vbNewLine & update_date
        EMWriteScreen "X", 6, 26
        transmit
        EMReadScreen cash_retro_total_income, 8, 9, 43
        EMReadScreen cash_retro_expenses, 8, 15, 43
        EMReadScreen cash_prosp_total_income, 8, 9, 59
        EMReadScreen cash_prosp_expenses, 8, 15, 59
        EMReadScreen cash_income_verif, 1, 9, 73
        EMReadScreen cash_expense_verif, 1, 15, 73

        EMReadScreen snap_retro_total_income, 8,  11, 43
        EMReadScreen snap_retro_expenses, 8, 17, 43
        EMReadScreen snap_prosp_total_income, 8,  11, 59
        EMReadScreen snap_prosp_expenses, 8, 17, 59
        EMReadScreen snap_income_verif, 1, 11, 73
        EMReadScreen snap_expense_verif, 1, 17, 73

        EMReadScreen hc_method_a_total_income, 8, 12, 59
        EMReadScreen hc_method_a_income_verif, 1, 12, 73
        EMReadScreen hc_method_a_expenses, 8, 18, 59
        EMReadScreen hc_method_a_expense_verif, 1, 18, 73

        EMReadScreen hc_method_b_total_income, 8, 13, 59
        EMReadScreen hc_method_b_income_verif, 1, 13, 73
        EMReadScreen hc_method_b_expenses, 8, 19, 59
        EMReadScreen hc_method_b_expense_verif, 1, 19, 73

        PF3

        If type_of_income = "01" Then busi_type = "01 - Farming"
        If type_of_income = "02" Then busi_type = "02 - Real Estate"
        If type_of_income = "03" Then busi_type = "03 - Home Product Sales"
        If type_of_income = "04" Then busi_type = "04 - Other Sales"
        If type_of_income = "05" Then busi_type = "05 - Personal Services"
        If type_of_income = "06" Then busi_type = "06 - Paper Route"
        If type_of_income = "07" Then busi_type = "07 - In Home Daycare"
        If type_of_income = "08" Then busi_type = "08 - Rental Income"
        If type_of_income = "09" Then busi_type = "09 - Other"

        income_start_date = replace(income_start_date, " ", "/")
        income_end_date = replace(income_end_date, " ", "/")
        If income_end_date = "__/__/__" Then income_end_date = ""
        ' cash_net_prosp_amount = trim(cash_net_prosp_amount)
        ' snap_net_prosp_amount = trim(snap_net_prosp_amount)
        ' hc_method_a_net_prosp_amount = trim(hc_method_a_net_prosp_amount)
        ' hc_method_b_net_prosp_amount = trim(hc_method_b_net_prosp_amount)
        ' reported_hours = trim(reported_hours)
        ' minimum_wage_hours = trim(minimum_wage_hours)
        update_date = replace(update_date, " ", "/")

        If SE_method = "01" Then SE_method = "01 - 50% Gross Inc"
        If SE_method = "02" Then SE_method = "02 - Tax Forms"

        SE_method_date = replace(SE_method_date, " ", "/")
        If SE_method_date = "__/__/__" Then SE_method_date = ""

        ' MsgBox "Line 898" & vbNewLine & update_date
        If cash_income_verif = "1" Then cash_verif = "1 - Tax Returns"
        If cash_income_verif = "2" Then cash_verif = "2 - Receipts"
        If cash_income_verif = "3" Then cash_verif = "3 - Busi Records"
        If cash_income_verif = "6" Then cash_verif = "6 - Other Doc"
        If cash_income_verif = "N" Then cash_verif = "N - No Verif Prvd"
        If cash_income_verif = "?" Then cash_verif = "? - Delayed Verif"

        If snap_income_verif = "1" Then snap_verif = "1 - Tax Returns"
        If snap_income_verif = "2" Then snap_verif = "2 - Receipts"
        If snap_income_verif = "3" Then snap_verif = "3 - Busi Records"
        If snap_income_verif = "4" Then snap_verif = "4 - Pend Out State"
        If snap_income_verif = "6" Then snap_verif = "6 - Other Doc"
        If snap_income_verif = "N" Then snap_verif = "N - No Verif Prvd"
        If snap_income_verif = "?" Then snap_verif = "? - Delayed Verif"

        If hc_method_b_income_verif = "1" Then hc_verif = "1 - Tax Returns"
        If hc_method_b_income_verif = "2" Then hc_verif = "2 - Receipts"
        If hc_method_b_income_verif = "3" Then hc_verif = "3 - Busi Records"
        If hc_method_b_income_verif = "6" Then hc_verif = "6 - Other Doc"
        If hc_method_b_income_verif = "N" Then hc_verif = "N - No Verif Prvd"
        If hc_method_b_income_verif = "?" Then hc_verif = "? - Delayed Verif"

        cash_retro_total_income = replace(cash_retro_total_income, "_", " ")
        ' cash_retro_total_income = trim(cash_retro_total_income)

        cash_retro_expenses = replace(cash_retro_expenses, "_", " ")
        ' cash_retro_expenses = trim(cash_retro_expenses)

        cash_prosp_total_income = replace(cash_prosp_total_income, "_", " ")
        ' cash_prosp_total_income = trim(cash_prosp_total_income)

        cash_prosp_expenses = replace(cash_prosp_expenses, "_", " ")
        ' cash_prosp_expenses = trim(cash_prosp_expenses)

        If cash_income_verif = "1" Then cash_income_verif = "1 - Tax Returns"
        If cash_income_verif = "2" Then cash_income_verif = "2 - Receipts"
        If cash_income_verif = "3" Then cash_income_verif = "3 - Busi Records"
        If cash_income_verif = "6" Then cash_income_verif = "6 - Other Doc"
        If cash_income_verif = "N" Then cash_income_verif = "N - No Verif Prvd"
        If cash_income_verif = "?" Then cash_income_verif = "? - Delayed Verif"

        If cash_expense_verif = "1" Then cash_expense_verif = "1 - Tax Returns"
        If cash_expense_verif = "2" Then cash_expense_verif = "2 - Receipts"
        If cash_expense_verif = "3" Then cash_expense_verif = "3 - Busi Records"
        If cash_expense_verif = "6" Then cash_expense_verif = "6 - Other Doc"
        If cash_expense_verif = "N" Then cash_expense_verif = "N - No Verif Prvd"
        If cash_expense_verif = "?" Then cash_expense_verif = "? - Delayed Verif"


        snap_retro_total_income = replace(snap_retro_total_income, "_", " ")
        ' snap_retro_total_income = trim(snap_retro_total_income)

        snap_retro_expenses = replace(snap_retro_expenses, "_", " ")
        ' snap_retro_expenses = trim(snap_retro_expenses)

        snap_prosp_total_income = replace(snap_prosp_total_income, "_", " ")
        ' snap_prosp_total_income = trim(snap_prosp_total_income)

        snap_prosp_expenses = replace(snap_prosp_expenses, "_", " ")
        ' snap_prosp_expenses = trim(snap_prosp_expenses)

        If snap_income_verif = "1" Then snap_income_verif = "1 - Tax Returns"
        If snap_income_verif = "2" Then snap_income_verif = "2 - Receipts"
        If snap_income_verif = "3" Then snap_income_verif = "3 - Busi Records"
        If snap_income_verif = "4" Then snap_income_verif = "4 - Pend Out State"
        If snap_income_verif = "6" Then snap_income_verif = "6 - Other Doc"
        If snap_income_verif = "N" Then snap_income_verif = "N - No Verif Prvd"
        If snap_income_verif = "?" Then snap_income_verif = "? - Delayed Verif"

        If snap_expense_verif = "1" Then snap_expense_verif = "1 - Tax Returns"
        If snap_expense_verif = "2" Then snap_expense_verif = "2 - Receipts"
        If snap_expense_verif = "3" Then snap_expense_verif = "3 - Busi Records"
        If snap_expense_verif = "4" Then snap_expense_verif = "4 - Pend Out State"
        If snap_expense_verif = "6" Then snap_expense_verif = "6 - Other Doc"
        If snap_expense_verif = "N" Then snap_expense_verif = "N - No Verif Prvd"
        If snap_expense_verif = "?" Then snap_expense_verif = "? - Delayed Verif"

        hc_method_a_total_income = replace(hc_method_a_total_income, "_", " ")
        ' hc_method_a_total_income = trim(hc_method_a_total_income)

        If hc_method_a_income_verif = "1" Then hc_method_a_income_verif = "1 - Tax Returns"
        If hc_method_a_income_verif = "2" Then hc_method_a_income_verif = "2 - Receipts"
        If hc_method_a_income_verif = "3" Then hc_method_a_income_verif = "3 - Busi Records"
        If hc_method_a_income_verif = "6" Then hc_method_a_income_verif = "6 - Other Doc"
        If hc_method_a_income_verif = "N" Then hc_method_a_income_verif = "N - No Verif Prvd"
        If hc_method_a_income_verif = "?" Then hc_method_a_income_verif = "? - Delayed Verif"

        hc_method_a_expenses = replace(hc_method_a_expenses, "_", " ")
        ' hc_method_a_expenses = trim(hc_method_a_expenses)

        If hc_method_a_expense_verif = "1" Then hc_method_a_expense_verif = "1 - Tax Returns"
        If hc_method_a_expense_verif = "2" Then hc_method_a_expense_verif = "2 - Receipts"
        If hc_method_a_expense_verif = "3" Then hc_method_a_expense_verif = "3 - Busi Records"
        If hc_method_a_expense_verif = "6" Then hc_method_a_expense_verif = "6 - Other Doc"
        If hc_method_a_expense_verif = "N" Then hc_method_a_expense_verif = "N - No Verif Prvd"
        If hc_method_a_expense_verif = "?" Then hc_method_a_expense_verif = "? - Delayed Verif"

        hc_method_b_total_income = replace(hc_method_b_total_income, "_", " ")
        ' hc_method_b_total_income = trim(hc_method_b_total_income)

        If hc_method_b_income_verif = "1" Then hc_method_b_income_verif = "1 - Tax Returns"
        If hc_method_b_income_verif = "2" Then hc_method_b_income_verif = "2 - Receipts"
        If hc_method_b_income_verif = "3" Then hc_method_b_income_verif = "3 - Busi Records"
        If hc_method_b_income_verif = "6" Then hc_method_b_income_verif = "6 - Other Doc"
        If hc_method_b_income_verif = "N" Then hc_method_b_income_verif = "N - No Verif Prvd"
        If hc_method_b_income_verif = "?" Then hc_method_b_income_verif = "? - Delayed Verif"

        hc_method_b_expenses = replace(hc_method_b_expenses, "_", " ")
        ' hc_method_b_expenses = trim(hc_method_b_expenses)

        If hc_method_b_expense_verif = "1" Then hc_method_b_expense_verif = "1 - Tax Returns"
        If hc_method_b_expense_verif = "2" Then hc_method_b_expense_verif = "2 - Receipts"
        If hc_method_b_expense_verif = "3" Then hc_method_b_expense_verif = "3 - Busi Records"
        If hc_method_b_expense_verif = "6" Then hc_method_b_expense_verif = "6 - Other Doc"
        If hc_method_b_expense_verif = "N" Then hc_method_b_expense_verif = "N - No Verif Prvd"
        If hc_method_b_expense_verif = "?" Then hc_method_b_expense_verif = "? - Delayed Verif"

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_UNEA_panel(access_type, member_name, unea_type, unea_verif, panel_claim_nmbr, start_date, end_date, cola_amt, unea_amount, unea_pay_amount, unea_frequency, update_date, panel_ref_numb, pic_ave_inc, pic_prosp_income, retro_total)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen member_name, 2, 4, 33
        EMReadScreen panel_type, 2, 5, 37
        EMReadScreen panel_verif_code, 1, 5, 65
        EMReadScreen panel_claim_nmbr, 15, 6, 37
        EMReadScreen panel_start_date, 8, 7, 37
        EMReadScreen panel_end_date, 8, 7, 68
        EMReadScreen cola_disregard, 8, 10, 67
        EMReadScreen update_date, 8, 21, 55

        For unea_row = 17 to 13 Step -1
            EMReadScreen pay_amount, 8, unea_row, 67
            If pay_amount <> "________" Then
                unea_pay_amount = trim(pay_amount)
                Exit For
            End If
        Next
        EMReadScreen total_amount, 8, 18, 68
        EMReadScreen retro_total, 8, 18, 39

        unea_amount = trim(total_amount)
        retro_total = trim(retro_total)

        EMWriteScreen "X", 10, 26       'opening SNAP pic
        transmit
        EMReadScreen pic_ave_inc, 8, 17, 56
        EMReadScreen pic_prosp_income, 8, 18, 56
        EMReadScreen panel_frequency_code, 1, 5, 64
        PF3

        If panel_type = "01" Then unea_type = "01 - RSDI, Disa"
        If panel_type = "02" Then unea_type = "02 - RSDI, No Disa"
        If panel_type = "03" Then unea_type = "03 - SSI"
        If panel_type = "06" Then unea_type = "06 - Non-MN PA"
        If panel_type = "11" Then unea_type = "11 - VA Disability"
        If panel_type = "12" Then unea_type = "12 - VA Pension"
        If panel_type = "13" Then unea_type = "13 - VA Other"
        If panel_type = "38" Then unea_type = "38 - VA Aid & Attendance"
        If panel_type = "14" Then unea_type = "14 - Unemployment Insurance"
        If panel_type = "15" Then unea_type = "15 - Worker's Comp"
        If panel_type = "16" Then unea_type = "16 - Railroad Retirement"
        If panel_type = "17" Then unea_type = "17 - Other Retirement"
        If panel_type = "18" Then unea_type = "18 - Military Enrirlement"
        If panel_type = "19" Then unea_type = "19 - FC Child req FS"
        If panel_type = "20" Then unea_type = "20 - FC Child not req FS"
        If panel_type = "21" Then unea_type = "21 - FC Adult req FS"
        If panel_type = "22" Then unea_type = "22 - FC Adult not req FS"
        If panel_type = "23" Then unea_type = "23 - Dividends"
        If panel_type = "24" Then unea_type = "24 - Interest"
        If panel_type = "25" Then unea_type = "25 - Cnt gifts/prizes"
        If panel_type = "26" Then unea_type = "26 - Strike Benefits"
        If panel_type = "27" Then unea_type = "27 - Contract for Deed"
        If panel_type = "28" Then unea_type = "28 - Illegal Income"
        If panel_type = "29" Then unea_type = "29 - Other Countable"
        If panel_type = "30" Then unea_type = "30 - Infrequent"
        If panel_type = "31" Then unea_type = "31 - Other - FS Only"
        If panel_type = "08" Then unea_type = "08 - Direct Child Support"
        If panel_type = "35" Then unea_type = "35 - Direct Spousal Support"
        If panel_type = "36" Then unea_type = "36 - Disbursed Child Support"
        If panel_type = "37" Then unea_type = "37 - Disbursed Spousal Support"
        If panel_type = "39" Then unea_type = "39 - Disbursed CS Arrears"
        If panel_type = "40" Then unea_type = "40 - Disbursed Spsl Sup Arrears"
        If panel_type = "43" Then unea_type = "43 - Disbursed Excess CS"
        If panel_type = "44" Then unea_type = "44 - MSA - Excess Income for SSI"
        If panel_type = "47" Then unea_type = "47 - Tribal Income"
        If panel_type = "48" Then unea_type = "48 - Trust Income"
        If panel_type = "49" Then unea_type = "49 - Non-Recurring"

        If panel_verif_code = "1" Then unea_verif = "1 - Copy of Checks"
        If panel_verif_code = "2" Then unea_verif = "2 - Award Letters"
        If panel_verif_code = "3" Then unea_verif = "3 - System Initiated Verif"
        If panel_verif_code = "4" Then unea_verif = "4 - Coltrl Stmt"
        If panel_verif_code = "5" Then unea_verif = "5 - Pend Out State Verif"
        If panel_verif_code = "6" Then unea_verif = "6 - Other Document"
        If panel_verif_code = "7" Then unea_verif = "7 - Worker Initiated Verif"
        If panel_verif_code = "8" Then unea_verif = "8 - RI Stubs"
        If panel_verif_code = "N" Then unea_verif = "N - No Verif Prvd"
        If panel_verif_code = "?" Then unea_verif = "? - Delayed Verif"

        panel_claim_nmbr = replace(panel_claim_nmbr, "_", "")

        start_date = replace(panel_start_date, " ", "/")
        end_date = replace(panel_end_date, " ", "/")
        If end_date = "__/__/__" Then end_date = ""
        update_date = replace(update_date, " ", "/")
        cola_amt = trim(cola_disregard)
        If cola_amt = "________" Then cola_amt = ""

        If panel_frequency_code = "1" Then unea_frequency = "1 - Monthly"
        If panel_frequency_code = "2" Then unea_frequency = "2 - Semi Monthly"
        If panel_frequency_code = "3" Then unea_frequency = "3 - Biweekly"
        If panel_frequency_code = "4" Then unea_frequency = "4 - Weekly"
        pic_ave_inc = trim(pic_ave_inc)
        pic_prosp_income = trim(pic_prosp_income)

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_ACCT_panel(access_type, member_name, account_type, account_number, account_location, account_balance, account_verification, update_date, panel_ref_numb, balance_date, withdraw_penalty, withdraw_yn, withdraw_verif_code, count_cash, count_snap, count_hc, count_grh, count_ive, joint_own_yn, share_ratio, next_interest)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen member_name, 2, 4, 33
        EMReadScreen panel_type, 2, 6, 44
        EMReadScreen panel_number, 20, 7, 44
        EMReadScreen panel_name, 20, 8, 44
        EMReadScreen panel_balance, 8, 10, 46
        EMReadScreen panel_verif_code, 1, 10, 64
        EMReadScreen balance_date, 8, 11, 44
        EMReadScreen withdraw_penalty, 8, 12, 46
        EMReadScreen withdraw_yn, 1, 12, 64
        EMReadScreen withdraw_verif_code, 1, 12, 72
        EMReadScreen count_cash, 1, 14, 50
        EMReadScreen count_snap, 1, 14, 57
        EMReadScreen count_hc, 1, 14, 64
        EMReadScreen count_grh, 1, 14, 72
        EMReadScreen count_ive, 1, 14, 80
        EMReadScreen joint_own_yn, 1, 15, 44
        EMReadScreen share_ratio, 5, 15, 76
        EMReadScreen next_interest, 5, 17, 57
        EMReadScreen update_date, 8, 21, 55

        If panel_type = "SV" Then account_type = "SV - Savings"
        If panel_type = "CK" Then account_type = "CK - Checking"
        If panel_type = "CE" Then account_type = "CE - Certificate of Deposit"
        If panel_type = "MM" Then account_type = "MM - Money Market"
        If panel_type = "DC" Then account_type = "DC - Debit Card"
        If panel_type = "KO" Then account_type = "KO - Keogh Account"
        If panel_type = "FT" Then account_type = "FT - Fed Thrift Savings Plan"
        If panel_type = "SL" Then account_type = "SL - State & Local Govt"
        If panel_type = "RA" Then account_type = "RA - Employee Ret Annuities"
        If panel_type = "NP" Then account_type = "NP - Non-Profit Emmployee Ret"
        If panel_type = "IR" Then account_type = "IR - Indiv Ret Acct"
        If panel_type = "RH" Then account_type = "RH - Roth IRA"
        If panel_type = "FR" Then account_type = "FR - Ret Plan for Employers"
        If panel_type = "CT" Then account_type = "CT - Corp Ret Trust"
        If panel_type = "RT" Then account_type = "RT - Other Ret Fund"
        If panel_type = "QT" Then account_type = "QT - Qualified Tuition (529)"
        If panel_type = "CA" Then account_type = "CA - Coverdell SV (530)"
        If panel_type = "OE" Then account_type = "OE - Other Educational"
        If panel_type = "OT" Then account_type = "OT - Other"

        account_number = replace(panel_number, "_", "")
        account_location =  replace(panel_name, "_", "")
        account_balance = trim(panel_balance)

        If panel_verif_code = "1"  Then account_verification = "1 - Bank Statement"
        If panel_verif_code = "2"  Then account_verification = "2 - Agcy Ver Form"
        If panel_verif_code = "3"  Then account_verification = "3 - Coltrl Contact"
        If panel_verif_code = "5"  Then account_verification = "5 - Other Document"
        If panel_verif_code = "6"  Then account_verification = "6 - Personal Statement"
        If panel_verif_code = "N"  Then account_verification = "N - No Ver Prvd"

        balance_date = replace(balance_date, " ", "/")
        If balance_date = "__/__/__" Then balance_date = ""

        withdraw_penalty = replace(withdraw_penalty, "_", "")
        withdraw_penalty = trim(withdraw_penalty)
        withdraw_yn = replace(withdraw_yn, "_", "")
        If withdraw_verif_code = "1"  Then withdraw_verif_code = "1 - Bank Statement"
        If withdraw_verif_code = "2"  Then withdraw_verif_code = "2 - Agcy Ver Form"
        If withdraw_verif_code = "3"  Then withdraw_verif_code = "3 - Coltrl Contact"
        If withdraw_verif_code = "5"  Then withdraw_verif_code = "5 - Other Document"
        If withdraw_verif_code = "6"  Then withdraw_verif_code = "6 - Personal Statement"
        If withdraw_verif_code = "N"  Then withdraw_verif_code = "N - No Ver Prvd"

        count_cash = replace(count_cash, "_", "")
        count_snap = replace(count_snap, "_", "")
        count_hc = replace(count_hc, "_", "")
        count_grh = replace(count_grh, "_", "")
        count_ive = replace(count_ive, "_", "")

        share_ratio = replace(share_ratio, " ", "")

        next_interest = replace(next_interest, " ", "/")
        If next_interest = "__/__" Then next_interest = ""

        update_date = replace(update_date, " ", "/")

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_CARS_panel(access_type, member_name, cars_type, cars_year, cars_make, cars_model, cars_verif, update_date, panel_ref_numb, cars_trade_in, cars_loan, cars_source, cars_owed, cars_owed_verif_code, cars_owed_date, cars_use, cars_hc_benefit, cars_joint_yn, cars_share)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen member_name, 2, 4, 33
        EMReadScreen cars_type, 1, 6, 43
        EMReadScreen cars_year, 4, 8, 31
        EMReadScreen cars_make, 15, 8, 43
        EMReadScreen cars_model, 15, 8, 66
        EMReadScreen cars_trade_in, 8, 9, 45            'not output
        EMReadScreen cars_loan, 8, 9, 62                'not output
        EMReadScreen cars_source, 1, 9, 80              'not output
        EMReadScreen cars_verif_code, 1, 10, 60
        EMReadScreen cars_owed, 8, 12, 45               'not output
        EMReadScreen cars_owed_verif_code, 1, 12, 60    'not output
        EMReadScreen cars_owed_date, 8, 13, 43          'not output
        EMReadScreen cars_use, 1, 15, 43                'not output
        EMReadScreen cars_hc_benefit, 1, 15, 76         'not output
        EMReadScreen cars_joint_yn, 1, 16, 43           'not output
        EMReadScreen cars_share, 5, 16, 76              'not output
        EMReadScreen cars_update, 8, 21, 55

        If cars_type = "1" Then cars_type = "1 - Car"
        If cars_type = "2" Then cars_type = "2 - Truck"
        If cars_type = "3" Then cars_type = "3 - Van"
        If cars_type = "4" Then cars_type = "4 - Camper"
        If cars_type = "5" Then cars_type = "5 - Motorcycle"
        If cars_type = "6" Then cars_type = "6 - Trailer"
        If cars_type = "7" Then cars_type = "7 - Other"

        cars_make = replace(cars_make, "_", "")
        cars_model = replace(cars_model, "_", "")


        cars_trade_in = replace(cars_trade_in, "_", "")
        cars_trade_in = trim(cars_trade_in)

        cars_loan = replace(cars_loan, "_", "")
        cars_loan = trim(cars_loan)

        If cars_source = "1" Then cars_source = "1 - NADA"
        If cars_source = "2" Then cars_source = "2 - Appraisal Val"
        If cars_source = "3" Then cars_source = "3 - Client Stmt"
        If cars_source = "4" Then cars_source = "4 - Other Document"

        If cars_verif_code = "1" Then cars_verif = "1 - Title"
        If cars_verif_code = "2" Then cars_verif = "2 - License Reg"
        If cars_verif_code = "3" Then cars_verif = "3 - DMV"
        If cars_verif_code = "4" Then cars_verif = "4 - Purchase Agmt"
        If cars_verif_code = "5" Then cars_verif = "5 - Other Document"
        If cars_verif_code = "N" Then cars_verif = "N - No Ver Prvd"

        cars_owed = replace(cars_owed, "_", "")
        cars_owed = trim(cars_owed)

        If cars_owed_verif_code = "1" Then cars_owed_verif_code = "1 - Bank/Lending Inst Stmt"
        If cars_owed_verif_code = "2" Then cars_owed_verif_code = "2 - Private Lender Stmt"
        If cars_owed_verif_code = "3" Then cars_owed_verif_code = "3 - Other Document"
        If cars_owed_verif_code = "4" Then cars_owed_verif_code = "4 - Pend Out State Verif"
        If cars_owed_verif_code = "N" Then cars_owed_verif_code = "N - No Ver Prvd"

        cars_owed_date = replace(cars_owed_date, " ", "/")
        If cars_owed_date = "__/__/__" Then cars_owed_date = ""

        If cars_use = "1" Then cars_use = "1 - Primary Vehicle"
        If cars_use = "2" Then cars_use = "2 - Employment/Training Search"
        If cars_use = "3" Then cars_use = "3 - Disa Transportation"
        If cars_use = "4" Then cars_use = "4 - Income Producing"
        If cars_use = "5" Then cars_use = "5 - Used as Home"
        If cars_use = "7" Then cars_use = "7 - Unlicensed"
        If cars_use = "8" Then cars_use = "8 - Other Countable"
        If cars_use = "9" Then cars_use = "9 - Unavailable"
        If cars_use = "0" Then cars_use = "0 - Long Distance Employment Travel"
        If cars_use = "A" Then cars_use = "A - Carry Heating Fuel or Water"

        cars_hc_benefit = replace(cars_hc_benefit, "_", "")
        cars_joint_yn = replace(cars_joint_yn, "_", "")
        cars_share = replace(cars_share, " ", "")

        update_date = replace(cars_update, " ", "/")

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_FACI_panel(access_type, notes_on_faci, facility_name, facility_vendor_number, facility_type, facility_FS_elig, FS_facility_type, facility_waiver_type, facility_LTC_inelig_reason, facility_inelig_begin_date, facility_inelig_end_date, facility_anticipated_out_date, facility_GRH_plan_required, facility_GRH_plan_verif, facility_cty_app_place, facility_approval_cty_name, facility_GRH_DOC_amount, facility_GRH_postpay, facility_stay_one_rate, facility_stay_one_date_in, facility_stay_one_date_out, facility_stay_two_rate, facility_stay_two_date_in, facility_stay_two_date_out, facility_stay_three_rate, facility_stay_three_date_in, facility_stay_three_date_out, facility_stay_four_rate, facility_stay_four_date_in, facility_stay_four_date_out, facility_stay_five_rate, facility_stay_five_date_in, facility_stay_five_date_out)
    If access_type = "READ" Then
        EMReadScreen facility_name,                 30, 6, 43
        EMReadScreen facility_vendor_number,        8, 5, 43
        EMReadScreen facility_type,                 2, 7, 43
        EMReadScreen facility_FS_elig,              1, 8, 43
        EMReadScreen FS_facility_type,              1, 8, 71
        EMReadScreen facility_waiver_type,          2, 7, 71
        EMReadScreen facility_LTC_inelig_reason,    1,  9, 43
        EMReadScreen facility_inelig_begin_date,    10, 10, 52
        EMReadScreen facility_inelig_end_date,      10, 10, 71
        EMReadScreen facility_anticipated_out_date, 10, 9, 71

        facility_name = replace(facility_name, "_", "")
        If facility_type = "41" Then facility_type = "41 - NF-I"
        If facility_type = "42" Then facility_type = "42 - NF-II"
        If facility_type = "43" Then facility_type = "43 - ICF-DD"
        If facility_type = "44" Then facility_type = "44 - Short Stay In NF-I"
        If facility_type = "45" Then facility_type = "45 - Short Stay In NF-II"
        If facility_type = "46" Then facility_type = "46 - Short Stay in ICF-DD"
        If facility_type = "47" Then facility_type = "47 - RTC - Not IMD"
        If facility_type = "48" Then facility_type = "48 - Medical Hospital"
        If facility_type = "49" Then facility_type = "49 - MSOP"
        If facility_type = "50" Then facility_type = "50 - IMD/RTC"
        If facility_type = "51" Then facility_type = "51 - Rule 31 CD-IMD"
        If facility_type = "52" Then facility_type = "52 - Rule 36 MI-IMD"
        If facility_type = "53" Then facility_type = "53 - IMD Hospitals"
        If facility_type = "55" Then facility_type = "55 - Adult Foster Care/Rule 203"
        If facility_type = "56" Then facility_type = "56 - GRH (Not FC or Rule 36)"
        If facility_type = "57" Then facility_type = "57 - Rule 36 MI-Non-IMD"
        If facility_type = "60" Then facility_type = "60 - Non-GRH"
        If facility_type = "61" Then facility_type = "61 - Rule 31 CD-Non-IMD"
        If facility_type = "67" Then facility_type = "67 - Family Violence Shelter"
        If facility_type = "68" Then facility_type = "68 - County Correctional Facility"
        If facility_type = "69" Then facility_type = "69 - Non-Cty Adult Correctional"

        If FS_facility_type = "1" Then FS_facility_type = "1 - Fed Subsidized Housing for Elderly"
        If FS_facility_type = "2" Then FS_facility_type = "2 - Licensed Facility/Treatment Center - CD"
        If FS_facility_type = "3" Then FS_facility_type = "3 - Blind or Disabled RSDI/SSI Recipient"
        If FS_facility_type = "4" Then FS_facility_type = "4 - Family Violence Shelter"
        If FS_facility_type = "5" Then FS_facility_type = "5 - Temporary Shelter for Homeless"
        If FS_facility_type = "6" Then FS_facility_type = "6 - Not a facility by FS Definition"

        If facility_waiver_type = "01" Then facility_waiver_type = "01 - CADI"
        If facility_waiver_type = "02" Then facility_waiver_type = "02 - CAC"
        If facility_waiver_type = "03" Then facility_waiver_type = "03 - EW Single"
        If facility_waiver_type = "04" Then facility_waiver_type = "04 - EW Married"
        If facility_waiver_type = "05" Then facility_waiver_type = "05 - TBI"
        If facility_waiver_type = "06" Then facility_waiver_type = "06 - DD"
        If facility_waiver_type = "07" Then facility_waiver_type = "07 - ACS (Alt Care Services DD)"
        If facility_waiver_type = "08" Then facility_waiver_type = "08 - SISEW Single"
        If facility_waiver_type = "09" Then facility_waiver_type = "09 - SISEW Married"

        If facility_LTC_inelig_reason = "L" Then facility_LTC_inelig_reason = "L - This level of Care Not Required"
        If facility_LTC_inelig_reason = "N" Then facility_LTC_inelig_reason = "N - Not Pre-Screened"
        If facility_LTC_inelig_reason = "_" Then facility_LTC_inelig_reason = ""

        facility_inelig_begin_date = replace(facility_inelig_begin_date, " ", "/")
        If facility_inelig_begin_date = "__/__/____" Then facility_inelig_begin_date = ""
        facility_inelig_end_date = replace(facility_inelig_end_date, " ", "/")
        If facility_inelig_end_date = "__/__/____" Then facility_inelig_end_date = ""
        facility_anticipated_out_date = replace(facility_anticipated_out_date, " ", "/")
        If facility_anticipated_out_date = "__/__/____" Then facility_anticipated_out_date = ""

        EMReadScreen facility_GRH_plan_required,    1, 11, 52
        EMReadScreen facility_cty_app_place,        1, 12, 52
        EMReadScreen facility_GRH_plan_verif,       1, 11, 71
        EMReadScreen facility_approval_cty,         2, 12, 71
        EMReadScreen facility_GRH_DOC_amount,       8, 13, 45
        EMReadScreen facility_GRH_postpay,          1, 13, 71

        EMReadScreen facility_stay_one_rate,        1,  14, 34
        EMReadScreen facility_stay_one_date_in,     10, 14, 47
        EMReadScreen facility_stay_one_date_out,    10, 14, 71

        EMReadScreen facility_stay_two_rate,        1,  15, 34
        EMReadScreen facility_stay_two_date_in,     10, 15, 47
        EMReadScreen facility_stay_two_date_out,    10, 15, 71

        EMReadScreen facility_stay_three_rate,      1,  16, 34
        EMReadScreen facility_stay_three_date_in,   10, 16, 47
        EMReadScreen facility_stay_three_date_out,  10, 16, 71

        EMReadScreen facility_stay_four_rate,       1,  17, 34
        EMReadScreen facility_stay_four_date_in,    10, 17, 47
        EMReadScreen facility_stay_four_date_out,   10, 17, 71

        EMReadScreen facility_stay_five_rate,       1,  18, 34
        EMReadScreen facility_stay_five_date_in,    10, 18, 47
        EMReadScreen facility_stay_five_date_out,   10, 18, 71

        facility_GRH_plan_required = replace(facility_GRH_plan_required, "_", "")
        facility_GRH_plan_verif = replace(facility_GRH_plan_verif, "_", "")
        facility_cty_app_place = replace(facility_cty_app_place, "_", "")
        Call get_county_name_from_county_code(facility_approval_cty, facility_approval_cty_name, TRUE)
        facility_GRH_DOC_amount = replace(facility_GRH_DOC_amount, "_", "")
        facility_GRH_DOC_amount = trim(facility_GRH_DOC_amount)
        facility_GRH_postpay = replace(facility_GRH_postpay, "_", "")

        If facility_stay_one_rate = "1" Then facility_stay_one_rate = "Rate 1"
        If facility_stay_one_rate = "2" Then facility_stay_one_rate = "Rate 2"
        If facility_stay_one_rate = "3" Then facility_stay_one_rate = "Rate 3"
        If facility_stay_one_rate = "_" Then facility_stay_one_rate = "      "
        facility_stay_one_date_in = replace(facility_stay_one_date_in, " ", "/")
        If facility_stay_one_date_in = "__/__/____" Then facility_stay_one_date_in = ""
        facility_stay_one_date_out = replace(facility_stay_one_date_out, " ", "/")
        If facility_stay_one_date_out = "__/__/____" Then facility_stay_one_date_out = ""

        If facility_stay_two_rate = "1" Then facility_stay_two_rate = "Rate 1"
        If facility_stay_two_rate = "2" Then facility_stay_two_rate = "Rate 2"
        If facility_stay_two_rate = "3" Then facility_stay_two_rate = "Rate 3"
        If facility_stay_two_rate = "_" Then facility_stay_two_rate = "      "
        facility_stay_two_date_in = replace(facility_stay_two_date_in, " ", "/")
        If facility_stay_two_date_in = "__/__/____" Then facility_stay_two_date_in = ""
        facility_stay_two_date_out = replace(facility_stay_two_date_out, " ", "/")
        If facility_stay_two_date_out = "__/__/____" Then facility_stay_two_date_out = ""

        If facility_stay_three_rate = "1" Then facility_stay_three_rate = "Rate 1"
        If facility_stay_three_rate = "2" Then facility_stay_three_rate = "Rate 2"
        If facility_stay_three_rate = "3" Then facility_stay_three_rate = "Rate 3"
        If facility_stay_three_rate = "_" Then facility_stay_three_rate = "      "
        facility_stay_three_date_in = replace(facility_stay_three_date_in, " ", "/")
        If facility_stay_three_date_in = "__/__/____" Then facility_stay_three_date_in = ""
        facility_stay_three_date_out = replace(facility_stay_three_date_out, " ", "/")
        If facility_stay_three_date_out = "__/__/____" Then facility_stay_three_date_out = ""

        If facility_stay_four_rate = "1" Then facility_stay_four_rate = "Rate 1"
        If facility_stay_four_rate = "2" Then facility_stay_four_rate = "Rate 2"
        If facility_stay_four_rate = "3" Then facility_stay_four_rate = "Rate 3"
        If facility_stay_four_rate = "_" Then facility_stay_four_rate = "      "
        facility_stay_four_date_in = replace(facility_stay_four_date_in, " ", "/")
        If facility_stay_four_date_in = "__/__/____" Then facility_stay_four_date_in = ""
        facility_stay_four_date_out = replace(facility_stay_four_date_out, " ", "/")
        If facility_stay_four_date_out = "__/__/____" Then facility_stay_four_date_out = ""

        If facility_stay_five_rate = "1" Then facility_stay_five_rate = "Rate 1"
        If facility_stay_five_rate = "2" Then facility_stay_five_rate = "Rate 2"
        If facility_stay_five_rate = "3" Then facility_stay_five_rate = "Rate 3"
        If facility_stay_five_rate = "_" Then facility_stay_five_rate = "      "
        facility_stay_five_date_in = replace(facility_stay_five_date_in, " ", "/")
        If facility_stay_five_date_in = "__/__/____" Then facility_stay_five_date_in = ""
        facility_stay_five_date_out = replace(facility_stay_five_date_out, " ", "/")
        If facility_stay_five_date_out = "__/__/____" Then facility_stay_five_date_out = ""
    End If
end function

function access_SECU_panel(access_type, member_name, security_type, security_account_number, security_name, security_cash_value, security_verif, secu_update_date, panel_ref_numb, security_face_value, security_withdraw, security_withdraw_yn, security_withdraw_verif, secu_cash_yn, secu_snap_yn, secu_hc_yn, secu_grh_yn, secu_ive_yn, secu_joint, secu_ratio, security_eff_date)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen member_name, 2, 4, 33
        EMReadScreen panel_type, 2, 6, 50
        EMReadScreen security_account_number, 12, 7, 50
        EMReadScreen security_name, 20, 8, 50
        EMReadScreen security_cash_value, 8, 10, 52
        EMReadScreen security_eff_date, 8, 11, 35   'not output
        EMReadScreen verif_code, 1, 11, 50
        EMReadScreen security_face_value, 8, 12, 52     'not output
        EMReadScreen security_withdraw, 8, 13, 52       'not output
        EMReadScreen security_withdraw_yn, 1, 13, 72    'not output
        EMReadScreen security_withdraw_verif, 1, 13, 80 'not output

        EMReadScreen secu_cash_yn, 1, 15, 50    'not output
        EMReadScreen secu_snap_yn, 1, 15, 57    'not output
        EMReadScreen secu_hc_yn, 1, 15, 64      'not output
        EMReadScreen secu_grh_yn, 1, 15, 72     'not output
        EMReadScreen secu_ive_yn, 1, 15, 80     'not output

        EMReadScreen secu_joint, 1, 16, 44      'not output
        EMReadScreen secu_ratio, 5, 16, 76      'not output
        EMReadScreen secu_update_date, 8, 21, 55

        If panel_type = "LI" Then security_type = "LI - Life Insurance"
        If panel_type = "ST" Then security_type = "ST - Stocks"
        If panel_type = "BO" Then security_type = "BO - Bonds"
        If panel_type = "CD" Then security_type = "CD - Ctrct for Deed"
        If panel_type = "MO" Then security_type = "MO - Mortgage Note"
        If panel_type = "AN" Then security_type = "AN - Annuity"
        If panel_type = "OT" Then security_type = "OT - Other"

        security_account_number = replace(security_account_number, "_", "")
        security_name = replace(security_name, "_", "")

        security_cash_value = replace(security_cash_value, "_", "")
        security_cash_value = trim(security_cash_value)

        security_eff_date = replace(security_eff_date, " ", "/")
        If security_eff_date = "__/__/__" Then security_eff_date = ""

        If verif_code = "1" Then security_verif = "1 - Agency Form"
        If verif_code = "2" Then security_verif = "2 - Source Doc"
        If verif_code = "3" Then security_verif = "3 - Phone Contact"
        If verif_code = "5" Then security_verif = "5 - Other Document"
        If verif_code = "6" Then security_verif = "6 - Personal Statement"
        If verif_code = "N" Then security_verif = "N - No Ver Prov"

        security_face_value = replace(security_face_value, "_", "")
        security_face_value = trim(security_face_value)

        security_withdraw = replace(security_withdraw, "_", "")
        security_withdraw = trim(security_withdraw)

        security_withdraw_yn = replace(security_withdraw_yn, "_", "")

        If security_withdraw_verif = "1" Then security_withdraw_verif = "1 - Agency Form"
        If security_withdraw_verif = "2" Then security_withdraw_verif = "2 - Source Doc"
        If security_withdraw_verif = "3" Then security_withdraw_verif = "3 - Phone Contact"
        If security_withdraw_verif = "4" Then security_withdraw_verif = "4 - Other Document"
        If security_withdraw_verif = "5" Then security_withdraw_verif = "5 - Personal Stmt"
        If security_withdraw_verif = "N" Then security_withdraw_verif = "N - No Ver Prov"

        secu_cash_yn = replace(secu_cash_yn, "_", "")
        secu_snap_yn = replace(secu_snap_yn, "_", "")
        secu_hc_yn = replace(secu_hc_yn, "_", "")
        secu_grh_yn = replace(secu_grh_yn, "_", "")
        secu_ive_yn = replace(secu_ive_yn, "_", "")

        secu_joint = replace(secu_joint, "_", "")
        secu_ratio = replace(secu_ratio, " ", "")

        secu_update_date = replace(secu_update_date, " ", "/")

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_SHEL_panel(access_type, hud_sub_yn, shared_yn, paid_to, rent_retro_amt, rent_retro_verif, rent_prosp_amt, rent_prosp_verif, lot_rent_retro_amt, lot_rent_retro_verif, lot_rent_prosp_amt, lot_rent_prosp_verif, mortgage_retro_amt, mortgage_retro_verif, mortgage_prosp_amt, mortgage_prosp_verif, insurance_retro_amt, insurance_retro_verif, insurance_prosp_amt, insurance_prosp_verif, tax_retro_amt, tax_retro_verif, tax_prosp_amt, tax_prosp_verif, room_retro_amt, room_retro_verif, room_prosp_amt, room_prosp_verif, garage_retro_amt, garage_retro_verif, garage_prosp_amt, garage_prosp_verif, subsidy_retro_amt, subsidy_retro_verif, subsidy_prosp_amt, subsidy_prosp_verif)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen hud_sub_yn,            1, 6, 46
        EMReadScreen shared_yn,             1, 6, 64
        EMReadScreen paid_to,               25, 7, 50

        paid_to = replace(paid_to, "_", "")

        EMReadScreen rent_retro_amt,        8, 11, 37
        EMReadScreen rent_retro_verif,      2, 11, 48
        EMReadScreen rent_prosp_amt,        8, 11, 56
        EMReadScreen rent_prosp_verif,      2, 11, 67

        rent_retro_amt = replace(rent_retro_amt, "_", "")
        rent_retro_amt = trim(rent_retro_amt)
        If rent_retro_verif = "SF" Then rent_retro_verif = "SF - Shelter Form"
        If rent_retro_verif = "LE" Then rent_retro_verif = "LE - Lease"
        If rent_retro_verif = "RE" Then rent_retro_verif = "RE - Rent Receipt"
        If rent_retro_verif = "OT" Then rent_retro_verif = "OT - Other Document"
        If rent_retro_verif = "NC" Then rent_retro_verif = "NC - Chg Rept, Neg Impact"
        If rent_retro_verif = "PC" Then rent_retro_verif = "PC - Chg Rept, Pos Imact"
        If rent_retro_verif = "NO" Then rent_retro_verif = "NO - No Ver Prvd"
        If rent_retro_verif = "__" Then rent_retro_verif = ""
        rent_prosp_amt = replace(rent_prosp_amt, "_", "")
        rent_prosp_amt = trim(rent_prosp_amt)
        If rent_prosp_verif = "SF" Then rent_prosp_verif = "SF - Shelter Form"
        If rent_prosp_verif = "LE" Then rent_prosp_verif = "LE - Lease"
        If rent_prosp_verif = "RE" Then rent_prosp_verif = "RE - Rent Receipt"
        If rent_prosp_verif = "OT" Then rent_prosp_verif = "OT - Other Document"
        If rent_prosp_verif = "NC" Then rent_prosp_verif = "NC - Chg Rept, Neg Impact"
        If rent_prosp_verif = "PC" Then rent_prosp_verif = "PC - Chg Rept, Pos Imact"
        If rent_prosp_verif = "NO" Then rent_prosp_verif = "NO - No Ver Prvd"
        If rent_prosp_verif = "__" Then rent_prosp_verif = ""

        EMReadScreen lot_rent_retro_amt,    8, 12, 37
        EMReadScreen lot_rent_retro_verif,  2, 12, 48
        EMReadScreen lot_rent_prosp_amt,    8, 12, 56
        EMReadScreen lot_rent_prosp_verif,  2, 12, 67

        lot_rent_retro_amt = replace(lot_rent_retro_amt, "_", "")
        lot_rent_retro_amt = trim(lot_rent_retro_amt)
        If lot_rent_retro_verif = "LE" Then lot_rent_retro_verif = "LE - Lease"
        If lot_rent_retro_verif = "RE" Then lot_rent_retro_verif = "RE - Rent Receipt"
        If lot_rent_retro_verif = "BI" Then lot_rent_retro_verif = "BI - Billing Stmt"
        If lot_rent_retro_verif = "OT" Then lot_rent_retro_verif = "OT - Other Document"
        If lot_rent_retro_verif = "NC" Then lot_rent_retro_verif = "NC - Chg Rept, Neg Impact"
        If lot_rent_retro_verif = "PC" Then lot_rent_retro_verif = "PC - Chg Rept, Pos Imact"
        If lot_rent_retro_verif = "NO" Then lot_rent_retro_verif = "NO - No Ver Prvd"
        If lot_rent_retro_verif = "__" Then lot_rent_retro_verif = ""
        lot_rent_prosp_amt = replace(lot_rent_prosp_amt, "_", "")
        lot_rent_prosp_amt = trim(lot_rent_prosp_amt)
        If lot_rent_prosp_verif = "LE" Then lot_rent_prosp_verif = "LE - Lease"
        If lot_rent_prosp_verif = "RE" Then lot_rent_prosp_verif = "RE - Rent Receipt"
        If lot_rent_prosp_verif = "BI" Then lot_rent_prosp_verif = "BI - Billing Stmt"
        If lot_rent_prosp_verif = "OT" Then lot_rent_prosp_verif = "OT - Other Document"
        If lot_rent_prosp_verif = "NC" Then lot_rent_prosp_verif = "NC - Chg Rept, Neg Impact"
        If lot_rent_prosp_verif = "PC" Then lot_rent_prosp_verif = "PC - Chg Rept, Pos Imact"
        If lot_rent_prosp_verif = "NO" Then lot_rent_prosp_verif = "NO - No Ver Prvd"
        If lot_rent_prosp_verif = "__" Then lot_rent_prosp_verif = ""

        EMReadScreen mortgage_retro_amt,    8, 13, 37
        EMReadScreen mortgage_retro_verif,  2, 13, 48
        EMReadScreen mortgage_prosp_amt,    8, 13, 56
        EMReadScreen mortgage_prosp_verif,  2, 13, 67

        mortgage_retro_amt = replace(mortgage_retro_amt, "_", "")
        mortgage_retro_amt = trim(mortgage_retro_amt)
        If mortgage_retro_verif = "MO" Then mortgage_retro_verif = "MO - Mortgage Pmt Book"
        If mortgage_retro_verif = "CD" Then mortgage_retro_verif = "CD - Ctrct fro Deed"
        If mortgage_retro_verif = "OT" Then mortgage_retro_verif = "OT - Other Document"
        If mortgage_retro_verif = "NC" Then mortgage_retro_verif = "NC - Chg Rept, Neg Impact"
        If mortgage_retro_verif = "PC" Then mortgage_retro_verif = "PC - Chg Rept, Pos Imact"
        If mortgage_retro_verif = "NO" Then mortgage_retro_verif = "NO - No Ver Prvd"
        If mortgage_retro_verif = "__" Then mortgage_retro_verif = ""
        mortgage_prosp_amt = replace(mortgage_prosp_amt, "_", "")
        mortgage_prosp_amt = trim(mortgage_prosp_amt)
        If mortgage_prosp_verif = "MO" Then mortgage_prosp_verif = "MO - Mortgage Pmt Book"
        If mortgage_prosp_verif = "CD" Then mortgage_prosp_verif = "CD - Ctrct fro Deed"
        If mortgage_prosp_verif = "OT" Then mortgage_prosp_verif = "OT - Other Document"
        If mortgage_prosp_verif = "NC" Then mortgage_prosp_verif = "NC - Chg Rept, Neg Impact"
        If mortgage_prosp_verif = "PC" Then mortgage_prosp_verif = "PC - Chg Rept, Pos Imact"
        If mortgage_prosp_verif = "NO" Then mortgage_prosp_verif = "NO - No Ver Prvd"
        If mortgage_prosp_verif = "__" Then mortgage_prosp_verif = ""

        EMReadScreen insurance_retro_amt,   8, 14, 37
        EMReadScreen insurance_retro_verif, 2, 14, 48
        EMReadScreen insurance_prosp_amt,   8, 14, 56
        EMReadScreen insurance_prosp_verif, 2, 14, 67

        insurance_retro_amt = replace(insurance_retro_amt, "_", "")
        insurance_retro_amt = trim(insurance_retro_amt)
        If insurance_retro_verif = "BI" Then insurance_retro_verif = "BI - Billing Stmt"
        If insurance_retro_verif = "OT" Then insurance_retro_verif = "OT - Other Document"
        If insurance_retro_verif = "NC" Then insurance_retro_verif = "NC - Chg Rept, Neg Impact"
        If insurance_retro_verif = "PC" Then insurance_retro_verif = "PC - Chg Rept, Pos Imact"
        If insurance_retro_verif = "NO" Then insurance_retro_verif = "NO - No Ver Prvd"
        If insurance_retro_verif = "__" Then insurance_retro_verif = ""
        insurance_prosp_amt = replace(insurance_prosp_amt, "_", "")
        insurance_prosp_amt = trim(insurance_prosp_amt)
        If insurance_prosp_verif = "BI" Then insurance_prosp_verif = "BI - Billing Stmt"
        If insurance_prosp_verif = "OT" Then insurance_prosp_verif = "OT - Other Document"
        If insurance_prosp_verif = "NC" Then insurance_prosp_verif = "NC - Chg Rept, Neg Impact"
        If insurance_prosp_verif = "PC" Then insurance_prosp_verif = "PC - Chg Rept, Pos Imact"
        If insurance_prosp_verif = "NO" Then insurance_prosp_verif = "NO - No Ver Prvd"
        If insurance_prosp_verif = "__" Then insurance_prosp_verif = ""

        EMReadScreen tax_retro_amt,         8, 15, 37
        EMReadScreen tax_retro_verif,       2, 15, 48
        EMReadScreen tax_prosp_amt,         8, 15, 56
        EMReadScreen tax_prosp_verif,       2, 15, 67

        tax_retro_amt = replace(tax_retro_amt, "_", "")
        tax_retro_amt = trim(tax_retro_amt)
        If tax_retro_verif = "TX" Then tax_retro_verif = "TX - Prop Tax Stmt"
        If tax_retro_verif = "OT" Then tax_retro_verif = "OT - Other Document"
        If tax_retro_verif = "NC" Then tax_retro_verif = "NC - Chg Rept, Neg Impact"
        If tax_retro_verif = "PC" Then tax_retro_verif = "PC - Chg Rept, Pos Imact"
        If tax_retro_verif = "NO" Then tax_retro_verif = "NO - No Ver Prvd"
        If tax_retro_verif = "__" Then tax_retro_verif = ""
        tax_prosp_amt = replace(tax_prosp_amt, "_", "")
        tax_prosp_amt = trim(tax_prosp_amt)
        If tax_prosp_verif = "TX" Then tax_prosp_verif = "TX - Prop Tax Stmt"
        If tax_prosp_verif = "OT" Then tax_prosp_verif = "OT - Other Document"
        If tax_prosp_verif = "NC" Then tax_prosp_verif = "NC - Chg Rept, Neg Impact"
        If tax_prosp_verif = "PC" Then tax_prosp_verif = "PC - Chg Rept, Pos Imact"
        If tax_prosp_verif = "NO" Then tax_prosp_verif = "NO - No Ver Prvd"
        If tax_prosp_verif = "__" Then tax_prosp_verif = ""

        EMReadScreen room_retro_amt,        8, 16, 37
        EMReadScreen room_retro_verif,      2, 16, 48
        EMReadScreen room_prosp_amt,        8, 16, 56
        EMReadScreen room_prosp_verif,      2, 16, 67

        room_retro_amt = replace(room_retro_amt, "_", "")
        room_retro_amt = trim(room_retro_amt)
        If room_retro_verif = "SF" Then room_retro_verif = "SF - Shelter Form"
        If room_retro_verif = "LE" Then room_retro_verif = "LE - Lease"
        If room_retro_verif = "RE" Then room_retro_verif = "RE - Rent Receipt"
        If room_retro_verif = "OT" Then room_retro_verif = "OT - Other Document"
        If room_retro_verif = "NC" Then room_retro_verif = "NC - Chg Rept, Neg Impact"
        If room_retro_verif = "PC" Then room_retro_verif = "PC - Chg Rept, Pos Imact"
        If room_retro_verif = "NO" Then room_retro_verif = "NO - No Ver Prvd"
        If room_retro_verif = "__" Then room_retro_verif = ""
        room_prosp_amt = replace(room_prosp_amt, "_", "")
        room_prosp_amt = trim(room_prosp_amt)
        If room_prosp_verif = "SF" Then room_prosp_verif = "SF - Shelter Form"
        If room_prosp_verif = "LE" Then room_prosp_verif = "LE - Lease"
        If room_prosp_verif = "RE" Then room_prosp_verif = "RE - Rent Receipt"
        If room_prosp_verif = "OT" Then room_prosp_verif = "OT - Other Document"
        If room_prosp_verif = "NC" Then room_prosp_verif = "NC - Chg Rept, Neg Impact"
        If room_prosp_verif = "PC" Then room_prosp_verif = "PC - Chg Rept, Pos Imact"
        If room_prosp_verif = "NO" Then room_prosp_verif = "NO - No Ver Prvd"
        If room_prosp_verif = "__" Then room_prosp_verif = ""

        EMReadScreen garage_retro_amt,      8, 17, 37
        EMReadScreen garage_retro_verif,    2, 17, 48
        EMReadScreen garage_prosp_amt,      8, 17, 56
        EMReadScreen garage_prosp_verif,    2, 17, 67

        garage_retro_amt = replace(garage_retro_amt, "_", "")
        garage_retro_amt = trim(garage_retro_amt)
        If garage_retro_verif = "SF" Then garage_retro_verif = "SF - Shelter Form"
        If garage_retro_verif = "LE" Then garage_retro_verif = "LE - Lease"
        If garage_retro_verif = "RE" Then garage_retro_verif = "RE - Rent Receipt"
        If garage_retro_verif = "OT" Then garage_retro_verif = "OT - Other Document"
        If garage_retro_verif = "NC" Then garage_retro_verif = "NC - Chg Rept, Neg Impact"
        If garage_retro_verif = "PC" Then garage_retro_verif = "PC - Chg Rept, Pos Imact"
        If garage_retro_verif = "NO" Then garage_retro_verif = "NO - No Ver Prvd"
        If garage_retro_verif = "__" Then garage_retro_verif = ""
        garage_prosp_amt = replace(garage_prosp_amt, "_", "")
        garage_prosp_amt = trim(garage_prosp_amt)
        If garage_prosp_verif = "SF" Then garage_prosp_verif = "SF - Shelter Form"
        If garage_prosp_verif = "LE" Then garage_prosp_verif = "LE - Lease"
        If garage_prosp_verif = "RE" Then garage_prosp_verif = "RE - Rent Receipt"
        If garage_prosp_verif = "OT" Then garage_prosp_verif = "OT - Other Document"
        If garage_prosp_verif = "NC" Then garage_prosp_verif = "NC - Chg Rept, Neg Impact"
        If garage_prosp_verif = "PC" Then garage_prosp_verif = "PC - Chg Rept, Pos Imact"
        If garage_prosp_verif = "NO" Then garage_prosp_verif = "NO - No Ver Prvd"
        If garage_prosp_verif = "__" Then garage_prosp_verif = ""

        EMReadScreen subsidy_retro_amt,     8, 18, 37
        EMReadScreen subsidy_retro_verif,   2, 18, 48
        EMReadScreen subsidy_prosp_amt,     8, 18, 56
        EMReadScreen subsidy_prosp_verif,   2, 18, 67

        subsidy_retro_amt = replace(subsidy_retro_amt, "_", "")
        subsidy_retro_amt = trim(subsidy_retro_amt)
        If subsidy_retro_verif = "SF" Then subsidy_retro_verif = "SF - Shelter Form"
        If subsidy_retro_verif = "LE" Then subsidy_retro_verif = "LE - Lease"
        If subsidy_retro_verif = "OT" Then subsidy_retro_verif = "OT - Other Document"
        If subsidy_retro_verif = "NO" Then subsidy_retro_verif = "NO - No Ver Prvd"
        If subsidy_retro_verif = "__" Then subsidy_retro_verif = ""
        subsidy_prosp_amt = replace(subsidy_prosp_amt, "_", "")
        subsidy_prosp_amt = trim(subsidy_prosp_amt)
        If subsidy_prosp_verif = "SF" Then subsidy_prosp_verif = "SF - Shelter Form"
        If subsidy_prosp_verif = "LE" Then subsidy_prosp_verif = "LE - Lease"
        If subsidy_prosp_verif = "OT" Then subsidy_prosp_verif = "OT - Other Document"
        If subsidy_prosp_verif = "NO" Then subsidy_prosp_verif = "NO - No Ver Prvd"
        If subsidy_prosp_verif = "__" Then subsidy_prosp_verif = ""
    End If
end function

function access_REST_panel(access_type, member_name, rest_type, rest_verif, rest_update_date, panel_ref_numb, rest_market_value, value_verif_code, rest_amt_owed, amt_owed_verif_code, rest_eff_date, rest_status, rest_joint_yn, rest_ratio, repymt_agree_date)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen member_name, 2, 4, 33
        EMReadScreen type_code, 1, 6, 39
        EMReadScreen type_verif_code, 2, 6, 62
        EMReadScreen rest_market_value, 10, 8, 41
        EMReadScreen value_verif_code, 2, 8, 62
        EMReadScreen rest_amt_owed, 10, 9, 41
        EMReadScreen amt_owed_verif_code, 2, 9, 62
        EMReadScreen rest_eff_date, 8, 10, 39
        EMReadScreen rest_status, 1, 12, 54
        EMReadScreen rest_joint_yn, 1, 13, 54
        EMReadScreen rest_ratio, 5, 14, 54
        EMReadScreen repymt_agree_date, 8, 16, 62
        EMReadScreen rest_update_date, 8, 21, 55

        If type_code = "1" Then rest_type = "1 - House"
        If type_code = "2" Then rest_type = "2 - Land"
        If type_code = "3" Then rest_type = "3 - Buildings"
        If type_code = "4" Then rest_type = "4 - Mobile Home"
        If type_code = "5" Then rest_type = "5 - Life Estate"
        If type_code = "6" Then rest_type = "6 - Other"

        If type_verif_code = "TX" Then rest_verif = "TX - Property Tax Statement"
        If type_verif_code = "PU" Then rest_verif = "PU - Purchase Agreement"
        If type_verif_code = "TI" Then rest_verif = "TI - Title/Deed"
        If type_verif_code = "CD" Then rest_verif = "CD - Contract for Deed"
        If type_verif_code = "CO" Then rest_verif = "CO - County Record"
        If type_verif_code = "OT" Then rest_verif = "OT - Other Document"
        If type_verif_code = "NO" Then rest_verif = "NO - No Ver Prvd"

        rest_market_value = replace(rest_market_value, "_", "")
        rest_market_value = trim(rest_market_value)

        If value_verif_code = "TX" Then value_verif_code = "TX - Property Tax Statement"
        If value_verif_code = "PU" Then value_verif_code = "PU - Purchase Agreement"
        If value_verif_code = "AP" Then value_verif_code = "AP - Appraisal"
        If value_verif_code = "CO" Then value_verif_code = "CO - County Record"
        If value_verif_code = "OT" Then value_verif_code = "OT - Other Document"
        If value_verif_code = "NO" Then value_verif_code = "NO - No Ver Prvd"

        rest_amt_owed = replace(rest_amt_owed, "_", "")
        rest_amt_owed = trim(rest_amt_owed)

        If amt_owed_verif_code = "MO" Then amt_owed_verif_code = "TI - Title/Deed"
        If amt_owed_verif_code = "LN" Then amt_owed_verif_code = "CD - Contract for Deed"
        If amt_owed_verif_code = "CD" Then amt_owed_verif_code = "CD - Contract for Deed"
        If amt_owed_verif_code = "OT" Then amt_owed_verif_code = "OT - Other Document"
        If amt_owed_verif_code = "NO" Then amt_owed_verif_code = "NO - No Ver Prvd"

        rest_eff_date = replace(rest_eff_date, " ", "/")
        If rest_eff_date = "__/__/__" Then rest_eff_date = ""

        If rest_status = "1" Then rest_status = "1 - Home Residence"
        If rest_status = "2" Then rest_status = "2 - For Sale, IV-E Rpymt Agmt"
        If rest_status = "3" Then rest_status = "3 - Joint Owner, Unavailable"
        If rest_status = "4" Then rest_status = "4 - Income Producing"
        If rest_status = "5" Then rest_status = "5 - Future Residence"
        If rest_status = "6" Then rest_status = "6 - Other"
        If rest_status = "7" Then rest_status = "7 - For Sale, Unavailable"

        rest_joint_yn = replace(rest_joint_yn, "_", "")
        rest_ratio = replace(rest_ratio, "_", "")

        repymt_agree_date = replace(repymt_agree_date, " ", "/")
        If repymt_agree_date = "__/__/__" Then repymt_agree_date = ""

        rest_update_date = replace(rest_update_date, " ", "/")

        EMReadScreen panel_ref_numb, 1, 2, 73
        panel_ref_numb = "0" & panel_ref_numb
    End If
end function

function access_HEST_panel(access_type, all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        Call navigate_to_MAXIS_screen("STAT", "HEST")

        hest_col = 40
        Do
            EMReadScreen pers_paying, 2, 6, hest_col
            If pers_paying <> "__" Then
                all_persons_paying = all_persons_paying & ", " & pers_paying
            Else
                exit do
            End If
            hest_col = hest_col + 3
        Loop until hest_col = 70
        If left(all_persons_paying, 1) = "," Then all_persons_paying = right(all_persons_paying, len(all_persons_paying) - 2)

        EMReadScreen choice_date, 8, 7, 40
        EMReadScreen actual_initial_exp, 8, 8, 61

        EMReadScreen retro_heat_ac_yn, 1, 13, 34
        EMReadScreen retro_heat_ac_units, 2, 13, 42
        EMReadScreen retro_heat_ac_amt, 6, 13, 49
        EMReadScreen retro_electric_yn, 1, 14, 34
        EMReadScreen retro_electric_units, 2, 14, 42
        EMReadScreen retro_electric_amt, 6, 14, 49
        EMReadScreen retro_phone_yn, 1, 15, 34
        EMReadScreen retro_phone_units, 2, 15, 42
        EMReadScreen retro_phone_amt, 6, 15, 49

        EMReadScreen prosp_heat_ac_yn, 1, 13, 60
        EMReadScreen prosp_heat_ac_units, 2, 13, 68
        EMReadScreen prosp_heat_ac_amt, 6, 13, 75
        EMReadScreen prosp_electric_yn, 1, 14, 60
        EMReadScreen prosp_electric_units, 2, 14, 68
        EMReadScreen prosp_electric_amt, 6, 14, 75
        EMReadScreen prosp_phone_yn, 1, 15, 60
        EMReadScreen prosp_phone_units, 2, 15, 68
        EMReadScreen prosp_phone_amt, 6, 15, 75

        choice_date = replace(choice_date, " ", "/")
        If choice_date = "__/__/__" Then choice_date = ""
        actual_initial_exp = trim(actual_initial_exp)
        actual_initial_exp = replace(actual_initial_exp, "_", "")

        retro_heat_ac_yn = replace(retro_heat_ac_yn, "_", "")
        retro_heat_ac_units = replace(retro_heat_ac_units, "_", "")
        retro_heat_ac_amt = trim(retro_heat_ac_amt)
        retro_electric_yn = replace(retro_electric_yn, "_", "")
        retro_electric_units = replace(retro_electric_units, "_", "")
        retro_electric_amt = trim(retro_electric_amt)
        retro_phone_yn = replace(retro_phone_yn, "_", "")
        retro_phone_units = replace(retro_phone_units, "_", "")
        retro_phone_amt = trim(retro_phone_amt)

        prosp_heat_ac_yn = replace(prosp_heat_ac_yn, "_", "")
        prosp_heat_ac_units = replace(prosp_heat_ac_units, "_", "")
        prosp_heat_ac_amt = trim(prosp_heat_ac_amt)
        If prosp_heat_ac_amt = "" Then prosp_heat_ac_amt = 0
        prosp_electric_yn = replace(prosp_electric_yn, "_", "")
        prosp_electric_units = replace(prosp_electric_units, "_", "")
        prosp_electric_amt = trim(prosp_electric_amt)
        If prosp_electric_amt = "" Then prosp_electric_amt = 0
        prosp_phone_yn = replace(prosp_phone_yn, "_", "")
        prosp_phone_units = replace(prosp_phone_units, "_", "")
        prosp_phone_amt = trim(prosp_phone_amt)
        If prosp_phone_amt = "" Then prosp_phone_amt = 0

        total_utility_expense = 0
        If prosp_heat_ac_yn = "Y" Then
            total_utility_expense =  prosp_heat_ac_amt
        ElseIf prosp_electric_yn = "Y" AND prosp_phone_yn = "Y" Then
            total_utility_expense =  prosp_electric_amt + prosp_phone_amt
        ElseIf prosp_electric_yn = "Y" Then
            total_utility_expense =  prosp_electric_amt
        Elseif prosp_phone_yn = "Y" Then
            total_utility_expense =  prosp_phone_amt
        End If

    End If
end function

function access_WREG_panel(access_type, notes_on_wreg, clt_fs_pwe, clt_wreg_status, clt_defer_fset, clt_orient_date, clt_sanc_begin_date, clt_numb_of_sanc, clt_sanc_reasons, clt_abawd_status, clt_banked_months, clt_GA_elig_basis, clt_GA_coop, abawd_counted_months, abawd_info_list, second_abawd_period, second_set_info_list)
    access_type = UCase(access_type)
    If access_type = "READ" Then
        EMReadScreen clt_fs_pwe, 1, 6, 68
        EMReadScreen clt_wreg_status, 2, 8, 50
        EMReadScreen clt_defer_fset, 1, 8, 80
        EMReadScreen clt_orient_date, 8, 9, 50
        EMReadScreen clt_sanc_begin_date, 8, 10, 50
        EMReadScreen clt_numb_of_sanc, 2, 11, 50
        EMReadScreen clt_sanc_reasons, 2, 12, 50
        EMReadScreen clt_abawd_status, 2, 13, 50
        EMReadScreen clt_banked_months, 1, 14, 50
        EMReadScreen clt_GA_elig_basis, 2, 15, 50
        EMReadScreen clt_GA_coop, 2, 15, 78

        EmWriteScreen "x", 13, 57
        transmit
        bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
        bene_yr_row = 10
        abawd_counted_months = 0
        abawd_info_list = ""
        second_abawd_period = 0
        second_set_info_list = ""
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
                EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                abawd_counted_months_string = counted_date_month & "/" & counted_date_year
                abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
                abawd_counted_months = abawd_counted_months + 1				'adding counted months
            END IF

            'declaring & splitting the abawd months array
            If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)

            'counting and checking for second set of ABAWD months
            IF is_counted_month = "Y" or is_counted_month = "N" THEN
                EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                second_abawd_period = second_abawd_period + 1				'adding counted months
                second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
            END IF

            'declaring & splitting the second set of abawd months array
            If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)

            bene_mo_col = bene_mo_col - 4
            IF bene_mo_col = 15 THEN
                bene_yr_row = bene_yr_row - 1
                bene_mo_col = 63
            END IF
            month_count = month_count + 1
        LOOP until month_count = 36
        PF3

        clt_fs_pwe = replace(clt_fs_pwe, "_", "")
        If clt_wreg_status = "03" Then clt_wreg_status = "03 - Unfit for Employment"
        If clt_wreg_status = "04" Then clt_wreg_status = "04 - Resp for Care of Incapacitated Person"
        If clt_wreg_status = "05" Then clt_wreg_status = "05 - Age 60 or Older"
        If clt_wreg_status = "06" Then clt_wreg_status = "06 - Under Age 16"
        If clt_wreg_status = "07" Then clt_wreg_status = "07 - Age 16-17, Living w/ Caregiver"
        If clt_wreg_status = "08" Then clt_wreg_status = "08 - Resp for Care of Child under 6"
        If clt_wreg_status = "09" Then clt_wreg_status = "09 - Empl 30 hrs/wk or Earnings of 30 hrs/wk"
        If clt_wreg_status = "10" Then clt_wreg_status = "10 - Matching Grant Participant"
        If clt_wreg_status = "11" Then clt_wreg_status = "11 - Receiving or Applied for UI"
        If clt_wreg_status = "12" Then clt_wreg_status = "12 - Enrolled in School, Training, or Higher Ed"
        If clt_wreg_status = "13" Then clt_wreg_status = "13 - Participating in CD Program"
        If clt_wreg_status = "14" Then clt_wreg_status = "14 - Receiving MFIP"
        If clt_wreg_status = "20" Then clt_wreg_status = "20 - Pending/Receiving DWP"
        If clt_wreg_status = "15" Then clt_wreg_status = "15 - Age 16-17, NOT Living w/ Caregiver"
        If clt_wreg_status = "16" Then clt_wreg_status = "16 - 50-59 Years Old"
        If clt_wreg_status = "17" Then clt_wreg_status = "17 - Receiving RCA or GA"
        If clt_wreg_status = "21" Then clt_wreg_status = "21 - Resp for Care of Child under 18"
        If clt_wreg_status = "30" Then clt_wreg_status = "30 - Mandatory FSET Participant"
        If clt_wreg_status = "02" Then clt_wreg_status = "02 - Fail to Cooperate with FSET"
        If clt_wreg_status = "33" Then clt_wreg_status = "33 - Non-Coop being Referred"

        clt_defer_fset = replace(clt_defer_fset, "_", "")
        clt_orient_date = replace(clt_orient_date, " ", "/")
        IF clt_orient_date = "__/__/__" Then clt_orient_date = ""

        clt_sanc_begin_date = replace(clt_sanc_begin_date, " ", "/")
        IF clt_sanc_begin_date = "__/01/__" Then clt_sanc_begin_date = ""
        IF clt_numb_of_sanc = "01" Then clt_numb_of_sanc = "1st Sanction"
        IF clt_numb_of_sanc = "02" Then clt_numb_of_sanc = "2nd Sanction"
        IF clt_numb_of_sanc = "03" Then clt_numb_of_sanc = "3rd Sanction"
        If clt_numb_of_sanc = "__" Then clt_numb_of_sanc = ""
        If clt_sanc_reasons = "01" Then clt_sanc_reasons = "01 - Attend Orientation"
        If clt_sanc_reasons = "02" Then clt_sanc_reasons = "02 - Develop Work Plan"
        If clt_sanc_reasons = "03" Then clt_sanc_reasons = "03 - Follow Work Plan"
        If clt_sanc_reasons = "__" Then clt_sanc_reasons = ""

        If clt_abawd_status = "01" Then clt_abawd_status = "01 - Work Reg Exempt"
        If clt_abawd_status = "02" Then clt_abawd_status = "02 - Under Age 18"
        If clt_abawd_status = "03" Then clt_abawd_status = "03 - Age 50 or Over"
        If clt_abawd_status = "04" Then clt_abawd_status = "04 - Caregiver of Minor Child"
        If clt_abawd_status = "05" Then clt_abawd_status = "05 - Pregnant"
        If clt_abawd_status = "06" Then clt_abawd_status = "06 - Employed Avg of 20 hrs/wk"
        If clt_abawd_status = "07" Then clt_abawd_status = "07 - Work Experience Participant"
        If clt_abawd_status = "08" Then clt_abawd_status = "08 - Other E&T Services"
        If clt_abawd_status = "09" Then clt_abawd_status = "09 - Resides in a Waivered Area"
        If clt_abawd_status = "10" Then clt_abawd_status = "10 - ABAWD Counted Month"
        If clt_abawd_status = "11" Then clt_abawd_status = "11 - 2nd-3rd Month Period of Elig"
        If clt_abawd_status = "12" Then clt_abawd_status = "12 - RCA or GA Recipient"
        If clt_abawd_status = "13" Then clt_abawd_status = "13 - ABAWD Banked Months"
        clt_banked_months = replace(clt_banked_months, "_", "")

        If clt_GA_elig_basis = "04" Then clt_GA_elig_basis = "04 - Permanent Ill or Incap"
        If clt_GA_elig_basis = "05" Then clt_GA_elig_basis = "05 - Temporary Ill or Incap"
        If clt_GA_elig_basis = "06" Then clt_GA_elig_basis = "06 - Care of Ill or Incap Memb"
        If clt_GA_elig_basis = "07" Then clt_GA_elig_basis = "07 - Requires Services in Residence"
        If clt_GA_elig_basis = "09" Then clt_GA_elig_basis = "09 - Mentally Ill or Dev Disa"
        If clt_GA_elig_basis = "10" Then clt_GA_elig_basis = "10 - SSI/RSDI Pending"
        If clt_GA_elig_basis = "11" Then clt_GA_elig_basis = "11 - Appealing SSI/RSDI Denial"
        If clt_GA_elig_basis = "12" Then clt_GA_elig_basis = "12 - Advanced Age"
        If clt_GA_elig_basis = "13" Then clt_GA_elig_basis = "13 - Learning Disability"
        If clt_GA_elig_basis = "17" Then clt_GA_elig_basis = "17 - Protect/Court Ordered"
        If clt_GA_elig_basis = "20" Then clt_GA_elig_basis = "20 - Age 16 or 17 SS Approval"
        If clt_GA_elig_basis = "25" Then clt_GA_elig_basis = "25 - Emancipated Minor"
        If clt_GA_elig_basis = "28" Then clt_GA_elig_basis = "28 - Unemployable"
        If clt_GA_elig_basis = "29" Then clt_GA_elig_basis = "29 - Displaced Hmkr (FT Student)"
        If clt_GA_elig_basis = "30" Then clt_GA_elig_basis = "30 - Minor w/ Adult Unrelated"
        If clt_GA_elig_basis = "32" Then clt_GA_elig_basis = "32 - Adult ESL/Adult HS"
        If clt_GA_elig_basis = "99" Then clt_GA_elig_basis = "99 - No Elig Basis"
        If clt_GA_elig_basis = "__" Then clt_GA_elig_basis = ""

        If clt_GA_coop = "01" Then clt_GA_coop = "01 - Cooperating"
        If clt_GA_coop = "03" Then clt_GA_coop = "03 - Failed to Coop"
        If clt_GA_coop = "__" Then clt_GA_coop = ""
    end If
end function

function create_array_of_all_panels(Array_Name)
    ' Do
    '     Call back_to_SELF
    '     Call navigate_to_MAXIS_screen("STAT", "SUMM")
    '     EMReadScreen summ_check, 4, 2, 46
    ' Loop until summ_check = "SUMM"
    Call navigate_to_MAXIS_screen("STAT", "PNLP")

    mx_row = 3
    counter = 1
    number_of_panels = 0
    Do
        EMReadScreen panel_name, 4, mx_row, 5
        If panel_name = "Case" Then panel_name = "    "

        If panel_name <> "    " Then
            EMReadScreen panel_ref, 2, mx_row, 10
            If panel_name <> last_panel Then
                counter = 1
            ElseIf panel_ref <> last_ref Then
                counter = 1
            End If
            ReDim Preserve Array_Name(panel_notes_const, number_of_panels)

            Array_Name(the_panel_const, number_of_panels) = panel_name
            Array_Name(panel_btn_const, number_of_panels) = 100 + number_of_panels

            one_per_case = FALSE
            one_per_person = FALSE
            multiple_per_case = FALSE
            multiple_per_person = FALSE

            Array_Name(one_per_case_const, number_of_panels) = FALSE
            Array_Name(multiple_per_case_const, number_of_panels) = FALSE
            Array_Name(one_per_person_const, number_of_panels) = FALSE
            Array_Name(multiple_per_person_const, number_of_panels) = FALSE
            Array_Name(show_this_panel, number_of_panels) = TRUE

            If panel_name = "ADDR" Then one_per_case = TRUE
            If panel_name = "HEST" Then one_per_case = TRUE
            If panel_name = "REVW" Then one_per_case = TRUE
            If panel_name = "AREP" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "ALTP" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "TYPE" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "PROG" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "HCRE" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "EATS" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "SIBL" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "DSTT" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "PACT" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "SWKR" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "MISC" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "RESI" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "BILS" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "BUDG" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "MMSA" Then
                one_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If

            If panel_name = "MEMB" Then one_per_person = TRUE
            If panel_name = "WREG" Then one_per_person = TRUE
            If panel_name = "CASH" Then one_per_person = TRUE
            If panel_name = "SHEL" Then one_per_person = TRUE
            If panel_name = "MEMI" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "ALIA" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "PARE" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "IMIG" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "SPON" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "ADME" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "REMO" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "DISA" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "PREG" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "STRK" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "STWK" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "SCHL" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "EMPS" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "STIN" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "STEC" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "PBEN" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "LUMP" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "TRAC" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "DCEX" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "WKEX" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "COEX" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "ACUT" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "PDED" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "FMED" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "MEDI" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "DIET" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "TIME" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "EMMA" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "HCMI" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "SANC" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "DFLN" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "MSUR" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "SSRT" Then
                one_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If

            If panel_name = "FCFC" Then
                multiple_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "FCPL" Then
                multiple_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "ABPS" Then
                multiple_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "INSA" Then
                multiple_per_case = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If

            If panel_name = "ACCT" Then multiple_per_person = TRUE
            If panel_name = "SECU" Then multiple_per_person = TRUE
            If panel_name = "CARS" Then multiple_per_person = TRUE
            If panel_name = "REST" Then multiple_per_person = TRUE
            If panel_name = "BUSI" Then multiple_per_person = TRUE
            If panel_name = "JOBS" Then multiple_per_person = TRUE
            If panel_name = "UNEA" Then multiple_per_person = TRUE
            If panel_name = "FACI" Then
                multiple_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = TRUE
            End If
            If panel_name = "OTHR" Then
                multiple_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "TRAN" Then
                multiple_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "RBIC" Then
                multiple_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "ACCI" Then
                multiple_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If
            If panel_name = "DISQ" Then
                multiple_per_person = TRUE
                Array_Name(show_this_panel, number_of_panels) = FALSE
            End If

            If one_per_case = TRUE Then Array_Name(one_per_case_const, number_of_panels) = TRUE
            If multiple_per_case = TRUE Then Array_Name(multiple_per_case_const, number_of_panels) = TRUE
            If one_per_person = TRUE Then Array_Name(one_per_person_const, number_of_panels) = TRUE
            If multiple_per_person = TRUE THen Array_Name(multiple_per_person_const, number_of_panels) = TRUE

            'Can only be one per case
            If one_per_case = TRUE Then

            'Not ref number specific but has a counter
            ElseIf multiple_per_case = TRUE Then
                Array_Name(the_inst_const, number_of_panels) = "0" & counter
                counter = counter + 1
            'Ref number specific but can only be one of each'
            ElseIf one_per_person = TRUE Then
                Array_Name(the_memb_const, number_of_panels) = panel_ref
            ElseIf multiple_per_person = TRUE Then
                Array_Name(the_memb_const, number_of_panels) = panel_ref
                Array_Name(the_inst_const, number_of_panels) = "0" & counter
                counter = counter + 1
            End If
            last_panel = panel_name
            last_ref = panel_ref
            number_of_panels = number_of_panels + 1
        End If

        mx_row = mx_row + 1
        If mx_row = 20 Then
            transmit
            mx_row = 3
            EMReadScreen check_for_end, 4, 2, 50
            EMReadScreen check_for_end_two, 4, 2, 46
        End If
    Loop until check_for_end = "SELF" or check_for_end_two = "WRAP"
    If check_for_end_two = "WRAP" Then transmit
end function

function csr_dlg_q_1()
	Do
		Do
			Do
				'This dialog reviews address and household composition
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 450, 230, "CSR Detail and Address"
				  GroupBox 5, 5, 415, 60, "SR Programs"
				  Text 15, 20, 155, 10, "Are you processing a SNAP six-month report?"
				  DropListBox 165, 15, 40, 45, " "+chr(9)+"Yes"+chr(9)+"No", snap_sr_yn
				  ' Text 210, 20, 25, 10, "status:"
				  ' DropListBox 235, 15, 90, 45, "Select One..."+chr(9)+"I - Incomplete"+chr(9)+"U - Complete and Updt Req"+chr(9)+"N - Not Rcvd"+chr(9)+"A - Approved"+chr(9)+"O - Override Autoclose"+chr(9)+"T - Terminated"+chr(9)+"D - Denied", curr_snap_sr_status
				  Text 220, 20, 40, 10, "Month/Year"
				  EditBox 260, 15, 20, 15, snap_sr_mo
				  EditBox 285, 15, 20, 15, snap_sr_yr

				  Text 15, 35, 155, 10, "Are you processing a HC six-month report?"
				  DropListBox 165, 30, 40, 45, " "+chr(9)+"Yes"+chr(9)+"No", hc_sr_yn
				  ' Text 210, 35, 25, 10, "status:"
				  ' DropListBox 235, 30, 90, 45, "Select One..."+chr(9)+"I - Incomplete"+chr(9)+"U - Complete and Updt Req"+chr(9)+"N - Not Rcvd"+chr(9)+"A - Approved"+chr(9)+"O - Override Autoclose"+chr(9)+"T - Terminated"+chr(9)+"D - Denied", curr_hc_sr_status
				  Text 220, 35, 40, 10, "Month/Year"
				  EditBox 260, 30, 20, 15, hc_sr_mo
				  EditBox 285, 30, 20, 15, hc_sr_yr

				  Text 15, 50, 155, 10, "Are you processing a GRH six-month report?"
				  DropListBox 165, 45, 40, 45, " "+chr(9)+"Yes"+chr(9)+"No", grh_sr_yn
				  ' Text 210, 50, 25, 10, "status:"
				  ' DropListBox 235, 45, 90, 45, "Select One..."+chr(9)+"I - Incomplete"+chr(9)+"U - Complete and Updt Req"+chr(9)+"N - Not Rcvd"+chr(9)+"A - Approved"+chr(9)+"O - Override Autoclose"+chr(9)+"T - Terminated"+chr(9)+"D - Denied", curr_grh_sr_status
				  Text 220, 50, 40, 10, "Month/Year"
				  EditBox 260, 45, 20, 15, grh_sr_mo
				  EditBox 285, 45, 20, 15, grh_sr_yr

				  GroupBox 5, 70, 415, 30, "CSR Name Information"
				  Text 15, 85, 200, 10, "Who is listed as the member on the CSR Form?"
				  ComboBox 210, 80, 150, 45, all_the_clients+chr(9)+"Person Information Missing", client_on_csr_form

				  DropListBox 295, 115, 145, 45, "Does the residence address match?"+chr(9)+"Yes - the addresses are the same."+chr(9)+"No - there is a difference."+chr(9)+"RESI Address not Provided"+chr(9)+"No - New Address Entered", residence_address_match_yn
				  DropListBox 295, 145, 145, 45, "Does the mailing address match?"+chr(9)+"Yes - the addresses are the same."+chr(9)+"No - there is a difference."+chr(9)+"MAIL Address not Provided"+chr(9)+"No - New Address Entered", mailing_address_match_yn
				  DropListBox 355, 170, 85, 45, "Select One..."+chr(9)+"Yes - Homeless"+chr(9)+"No", homeless_status
				  ' EditBox 355, 190, 60, 15, form_phone_number
				  If new_resi_addr_entered = FALSE AND new_mail_addr_entered = FALSE Then GroupBox 5, 105, 130, 105, "Current Case Address in MAXIS"
				  If new_resi_addr_entered = TRUE OR new_mail_addr_entered = TRUE Then GroupBox 5, 105, 130, 105, "Case Address Information"
				  ' Text 15, 115, 75, 10, "Residence Address:"
				  If new_resi_addr_entered = TRUE Then
					  Text 15, 115, 110, 10, "UPDATED Residence Address:"
					  Text 20, 125, 110, 10, new_resi_one
					  Text 20, 135, 110, 10, new_resi_city & ", " & new_resi_state & " " & new_resi_zip
				  Else
					  Text 15, 115, 75, 10, "Residence Address:"
					  Text 20, 125, 110, 10, resi_line_one
					  If resi_line_two = "" Then
					    Text 20, 135, 110, 10, resi_city & ", " & resi_state & " " & resi_zip
					  Else
					    Text 20, 135, 110, 10, resi_line_two
					    Text 20, 145, 110, 10, resi_city & ", " & resi_state & " " & resi_zip
					  End If
				  End If
				  If new_mail_addr_entered = TRUE Then
					  Text 15, 160, 110, 10, "UPDATED Mailing Address:"
					  Text 20, 170, 110, 10, new_mail_one
					  Text 20, 180, 110, 10, new_mail_city & ", " & new_mail_state & " " & new_mail_zip
				  Else
					  Text 15, 160, 75, 10, "Mailing Address:"
					  If mail_line_one = "" Then
					      Text 20, 170, 110, 10, "NO MAILING ADDRESS LISTED"
					  Else
					      Text 20, 170, 110, 10, mail_line_one
					      If mail_line_two = "" Then
					        Text 20, 180, 110, 10, mail_city & ", " & mail_state & " " & mail_zip
					      Else
					        Text 20, 180, 110, 10, mail_line_two
					        Text 20, 190, 110, 10, mail_city & ", " & mail_state & " " & mail_zip
					      End If
					  End If
				  End IF
				  GroupBox 140, 105, 305, 105, "Address on CSR"
				  Text 150, 115, 135, 25, "RESIDENCE ADDRESS MATCH? Does the Residence Address in MAXIS match the Residence Address reported on the CSR?"
				  Text 150, 145, 135, 25, "MAILING ADDRESS MATCH? Does the Mailing Address in MAXIS match the Mailing Address reported on the CSR?"
				  Text 150, 175, 170, 10, "Does the CSR Form indicate the client is homeless?"
				  ' Text 150, 195, 170, 10, "Phone number listed on the CSR:"

				  ButtonGroup ButtonPressed
				    OkButton 340, 210, 50, 15
				    CancelButton 395, 210, 50, 15
				EndDialog

			    err_msg = ""

			    dialog Dialog1
			    cancel_confirmation

			    program_indicated = FALSE
			    If snap_sr_yn = "Yes" Then
			        ' If snap_sr_status = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this case is for a SNAP Six-Month report indicate the status of the SR process (eg. N, I, or U)."
			        Call validate_footer_month_entry(snap_sr_mo, snap_sr_yr, err_msg, "* SNAP SR MONTH")
			        program_indicated = TRUE
			    End If
			    If hc_sr_yn = "Yes" Then
			        ' If hc_sr_status = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this case is for a HC Six-Month report indicate the status of the SR process (eg. N, I, or U)."
			        Call validate_footer_month_entry(hc_sr_mo, hc_sr_yr, err_msg, "* HC SR MONTH")
			        program_indicated = TRUE
			    End If
			    If grh_sr_yn = "Yes" Then
			        ' If grh_sr_status = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since this case is for a GRH Six-Month report indicate the status of the SR process (eg. N, I, or U)."
			        Call validate_footer_month_entry(grh_sr_mo, grh_sr_yr, err_msg, "* GRH SR MONTH")
			        program_indicated = TRUE
			    End If

			    If client_on_csr_form = "Select or Type" OR trim(client_on_csr_form) = "" Then err_msg = err_msg & vbNewLine & "* Indicate who is listed on the CSR form in the person infromation, or if this is blank, select that the person information is missing."

			    If program_indicated = FALSE Then err_msg = err_msg & vbNewLine & "* Select the program(s) that the CSR form is processing. (None of the programs are indicated to have an SR due.)"

			    If residence_address_match_yn = "Does the residence address match?" Then err_msg = err_msg & vbNewLine & "* Indicate information about the residence address provided on the CSR form."
			    If mailing_address_match_yn = "Does the mailing address match?" Then err_msg = err_msg & vbNewLine & "* Indicate information abobut the mailing address provided on the CSR form."
				If residence_address_match_yn = "No - New Address Entered" AND new_resi_addr_entered = FALSE Then err_msg = err_msg & vbNewLine & "* The option 'No - New Address Endered' for the residence address can only be updated by the script."
				If mailing_address_match_yn = "No - New Address Entered" AND new_mail_addr_entered = FALSE Then err_msg = err_msg & vbNewLine & "* The option 'No - New Address Endered' for the Mailing address can only be updated by the script."
			    If homeless_yn = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if the CSR form indicates the household is homeless or not."

			    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
			Loop until err_msg = ""

			If residence_address_match_yn = "No - there is a difference." Then call enter_new_residence_address

			If mailing_address_match_yn = "No - there is a difference." Then call enter_new_mailing_address
		Loop until mailing_address_match_yn <> "No - there is a difference." AND residence_address_match_yn <> "No - there is a difference."

		show_csr_dlg_q_1 = FALSE
		csr_dlg_q_1_cleared = TRUE
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
end function

function csr_dlg_q_2()
    Do

		dlg_width = 425
		If grh_sr = TRUE Then dlg_width = dlg_width + 50
		If hc_sr = TRUE Then dlg_width = dlg_width + 50
		If snap_sr = TRUE Then dlg_width = dlg_width + 50

		dlg_len = 80
		For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
		    dlg_len = dlg_len + 15
		Next

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, dlg_width, dlg_len, "CSR Household"
		  ' GroupBox 5, 180, 415, grp_len, "Household Comp"
		  Text 15, 10, 220, 10, "Q2. Has anyone moved in or out of your home in the past six months?"
		  DropListBox 240, 5, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", quest_two_move_in_out
		  x_pos = 430
		  y_pos = 25
		  Text 15, y_pos + 5, 275, 10, "Are there new household members that have been reported that are not listed here?"
		  DropListBox 295, y_pos, 150, 45, "Select One..."+chr(9)+"Yes - add another member"+chr(9)+"No - all member in MAXIS"+chr(9)+"New Members Have been Added", new_hh_memb_not_in_mx_yn
		  y_pos = y_pos + 20
		  Text 15, y_pos, 35, 10, "Member #"
		  Text 60, y_pos, 40, 10, "Last Name"
		  Text 130, y_pos, 40, 10, "First Name"
		  Text 205, y_pos, 15, 10, "Age"
		  Text 295, y_pos, 50, 10, "HH Moved Out"
		  Text 360, y_pos, 55, 10, "HH Moved In"
		  If grh_sr = TRUE Then
		      Text x_pos, y_pos, 20, 10, "GRH"
		      grh_col = x_pos
		      x_pos = x_pos + 50
		  End If
		  If hc_sr = TRUE Then
		      Text x_pos, y_pos, 20, 10, "HC"
		      hc_col = x_pos
		      x_pos = x_pos + 50
		  End If
		  If snap_sr = TRUE Then
		      Text x_pos, y_pos, 20, 10, "SNAP"
		      snap_col = x_pos
		      x_pos = x_pos + 50
		  End If
		  y_pos = y_pos + 20
		  For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
		    Text 20, y_pos, 15, 10, ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb)
		    Text 60, y_pos, 65, 10, ALL_CLIENTS_ARRAY(memb_last_name, known_memb)
		    Text 130, y_pos, 65, 10, ALL_CLIENTS_ARRAY(memb_first_name, known_memb)
		    Text 205, y_pos, 30, 10, ALL_CLIENTS_ARRAY(memb_age, known_memb)
		    CheckBox 305, y_pos, 25, 10, "Out", ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb)
		    CheckBox 370, y_pos, 25, 10, "In", ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb)
		    x_pos = 430
		    If grh_sr = TRUE Then Text grh_col, y_pos, 35, 10, ALL_CLIENTS_ARRAY(clt_grh_status, known_memb)
		    If hc_sr = TRUE Then Text hc_col, y_pos, 35, 10, ALL_CLIENTS_ARRAY(clt_hc_status, known_memb)
		    If snap_sr = TRUE Then Text snap_col, y_pos, 35, 10, ALL_CLIENTS_ARRAY(clt_snap_status, known_memb)
		    y_pos = y_pos + 15
		  Next
		  ' y_pos = y_pos + 25
		  ButtonGroup ButtonPressed
		    OkButton dlg_width - 110, y_pos, 50, 15
		    CancelButton dlg_width - 55, y_pos, 50, 15
		EndDialog

        err_msg = ""

        dialog Dialog1
        cancel_confirmation

        If quest_two_move_in_out = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the answer for Question 2 as provided on the CSR Form."
        If new_hh_memb_not_in_mx_yn = "Select One..." Then err_msg = err_msg & vbNewLine & "* Indicate if there are new members of the household that are not listed on this dialog."
		If new_hh_memb_not_in_mx_yn = "New Members Have been Added" AND new_memb_counter = 0 Then err_msg = err_msg & vbNewLine & "* No new members have been added during this script run. Select either 'Yes' or 'No'."
        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
	show_csr_dlg_q_2 = FALSE
	csr_dlg_q_2_cleared = TRUE

	If new_hh_memb_not_in_mx_yn = "Yes - add another member" Then
	    Do
	        ReDim Preserve NEW_MEMBERS_ARRAY(new_memb_notes, new_memb_counter)

	        BeginDialog Dialog1, 0, 0, 255, 210, "New HH Member"
	          EditBox 55, 35, 120, 15, NEW_MEMBERS_ARRAY(new_first_name, new_memb_counter)
	          EditBox 235, 35, 15, 15, NEW_MEMBERS_ARRAY(new_mid_initial, new_memb_counter)
	          EditBox 55, 55, 120, 15, NEW_MEMBERS_ARRAY(new_last_name, new_memb_counter)
	          EditBox 210, 55, 40, 15, NEW_MEMBERS_ARRAY(new_suffix, new_memb_counter)
	          EditBox 55, 75, 50, 15, NEW_MEMBERS_ARRAY(new_dob, new_memb_counter)
	          DropListBox 105, 95, 145, 45, "Select One..."+chr(9)+"01 - Applicant"+chr(9)+"02 - Spouse"+chr(9)+"03 - Child"+chr(9)+"04 - Parent"+chr(9)+"05 - Sibling"+chr(9)+"06 - Step Sibling"+chr(9)+"08 - Step Child"+chr(9)+"09 - Step Parent"+chr(9)+"10 - Aunt"+chr(9)+"11 - Uncle"+chr(9)+"12 - Niece"+chr(9)+"13 - Nephew"+chr(9)+"14 - Cousin"+chr(9)+"15 - Grandparent"+chr(9)+"16 - Grandchild"+chr(9)+"17 - Other Relative"+chr(9)+"18 - Legal Guardian"+chr(9)+"24 - Not Related"+chr(9)+"25 - Live-In Attendant"+chr(9)+"27 - Unknown", NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter)
	          CheckBox 35, 130, 20, 10, "HC", NEW_MEMBERS_ARRAY(new_ma_request, new_memb_counter)
	          CheckBox 65, 130, 30, 10, "SNAP", NEW_MEMBERS_ARRAY(new_fs_request, new_memb_counter)
	          CheckBox 100, 130, 30, 10, "GRH", NEW_MEMBERS_ARRAY(new_grh_request, new_memb_counter)
	          CheckBox 200, 115, 50, 10, "Moved In", NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter)
	          CheckBox 200, 130, 50, 10, "Moved Out", NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter)
	          EditBox 40, 150, 210, 15, NEW_MEMBERS_ARRAY(new_memb_notes, new_memb_counter)
	          ButtonGroup ButtonPressed
	            PushButton 145, 190, 50, 15, "Add Another", add_another_new_memb_btn
	            PushButton 200, 190, 50, 15, "No More", done_adding_new_memb_btn
	          Text 10, 10, 155, 20, "Enter any information about the new household member that has not been added to MAXIS."
	          Text 10, 40, 40, 10, "First Name"
	          Text 180, 40, 45, 10, "Middle Initial:"
	          Text 10, 60, 40, 10, "Last Name:"
	          Text 180, 60, 25, 10, "Suffix:"
	          Text 10, 80, 45, 10, "Date of Birth:"
	          Text 10, 100, 85, 10, "Relationship to Memb 01:"
	          GroupBox 10, 115, 165, 30, "Check any programs this Memb is requesting"
	          Text 10, 155, 25, 10, "Notes:"
	          Text 15, 175, 95, 25, "This script will not add this information to STAT, it will CASE:NOTE the information."
	        EndDialog

	        Dialog Dialog1
	        cancel_confirmation

			If ButtonPressed = -1 Then ButtonPressed = add_another_new_memb_btn
			If ButtonPressed = 0 Then ButtonPressed = done_adding_new_memb_btn
	        If ButtonPressed = add_another_new_memb_btn Then new_memb_counter = new_memb_counter + 1
	    Loop until ButtonPressed = done_adding_new_memb_btn
		new_hh_memb_not_in_mx_yn = "New Members Have been Added"

	    For new_hh_memb = 0 to UBound(NEW_MEMBERS_ARRAY, 2)
	        NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) = trim(NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb))
	        NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb) = trim(NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb))
	        If NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) = "" Then NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) = NEW_MEMBERS_ARRAY(new_first_name, new_hh_memb) & " " & NEW_MEMBERS_ARRAY(new_last_name, new_hh_memb)
	        If NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) <> "" Then NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) = NEW_MEMBERS_ARRAY(new_first_name, new_hh_memb) & " " & NEW_MEMBERS_ARRAY(new_mid_initial, new_hh_memb) & ". " & NEW_MEMBERS_ARRAY(new_last_name, new_hh_memb)
	        If NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb) <> "" Then NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) = NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb) & " " & NEW_MEMBERS_ARRAY(new_suffix, new_hh_memb)
	    Next
		new_memb_counter = new_memb_counter + 1
	End If
end function

function csr_dlg_q_4_7()
	Do
		dlg_len = 190
		q_4_grp_len = 15
		q_5_grp_len = 30
		q_6_grp_len = 30
		q_7_grp_len = 30
		For new_jobs_listed = 0 to UBound(NEW_EARNED_ARRAY, 2)
			If NEW_EARNED_ARRAY(earned_type, new_jobs_listed) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, new_jobs_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_5_grp_len = q_5_grp_len + 20
			End If
			If NEW_EARNED_ARRAY(earned_type, new_jobs_listed) = "JOBS"  AND NEW_EARNED_ARRAY(earned_prog_list, new_jobs_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_6_grp_len = q_6_grp_len + 20
			End If
		Next
		For new_unea_listed = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			If NEW_UNEARNED_ARRAY(unearned_type, new_unea_listed) = "UNEA"  AND NEW_UNEARNED_ARRAY(unearned_prog_list, new_unea_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_7_grp_len = q_7_grp_len + 20
			End If
		Next
		' If apply_for_ma = "Yes" Then
		'     dlg_len = dlg_len + (UBound(NEW_MA_REQUEST_ARRAY, 2) + 1) * 20
		'     q_4_grp_len = 35 + UBound(NEW_MA_REQUEST_ARRAY, 2) * 20
		' End If
		For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
			If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
				dlg_len = dlg_len + 20
				q_4_grp_len = q_4_grp_len + 20
			End If
		Next

		y_pos = 45

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 610, dlg_len, "MA CSR Income Questions"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 4 - 7:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 4 - 7 are completely blank.", all_questions_4_7_blank

		  GroupBox 15, 30, 585, q_4_grp_len, "Q4. Do you want to apply for MA for someone who is not getting coverage now?"
		  ' DropListBox 285, 25, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", apply_for_ma
		  ' CheckBox 430, 30, 75, 10, "Q4 Deailts left Blank", q_4_details_blank_checkbox
		  DropListBox 285, 25, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", apply_for_ma
		  ButtonGroup ButtonPressed
			PushButton 540, 30, 50, 10, "Add Another", add_memb_btn
		  ' If apply_for_ma = "Yes" Then
		  For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
			  If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
				  Text 35, y_pos + 5, 105, 10, "Select the Member requesting:"
				  ComboBox 145, y_pos, 195, 45, all_the_clients, NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If y_pos = 45 Then y_pos = y_pos + 5
		  ' End If

		  GroupBox 15, y_pos + 5, 585, q_5_grp_len, "Q5. Is anyone self-employed or does anyone expect to be self-employed?"
		  ' DropListBox 265, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_self_employed
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q5 Deailts left Blank", q_5_details_blank_checkbox
		  DropListBox 265, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_self_employed
		  y_pos = y_pos + 20

		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_busi_btn
		  first_busi= TRUE
		  For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
			  If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
				  If first_busi = TRUE then
					  Text 35, y_pos, 25, 10, "Name"
					  Text 155, y_pos, 55, 10, "Business Name"
					  Text 265, y_pos, 35, 10, "Start Date"
					  Text 325, y_pos, 50, 10, "Yearly Income"
					  y_pos = y_pos + 10
					  first_busi = FALSE
				  End If

				  ComboBox 35, y_pos, 110, 45, all_the_clients, NEW_EARNED_ARRAY(earned_client, each_busi)
				  EditBox 155, y_pos, 105, 15, NEW_EARNED_ARRAY(earned_source, each_busi)
				  EditBox 265, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_start_date, each_busi)
				  EditBox 325, y_pos, 60, 15, NEW_EARNED_ARRAY(earned_amount, each_busi)
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_INCOME_ARRAY(new_checkbox, each_busi)
				  ' CheckBox 530, y_pos, 40, 10, "Detail", ALL_INCOME_ARRAY(update_checkbox, each_busi)
				  y_pos = y_pos  + 20
			  End If
		  Next
		  If first_busi = TRUE Then
			  Text 35, y_pos, 400, 10, "CSR form - no BUSI information has been added."
			  y_pos = y_pos + 20
		  Else
			y_pos = y_pos + 10
		  End If

		  GroupBox 15, y_pos + 5, 585, q_6_grp_len, "Q6. Does anyone work or does anyone expect to start working?"
		  ' DropListBox 230, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_start_working
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q6 Deailts left Blank", q_6_details_blank_checkbox
		  DropListBox 230, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_start_working
		  y_pos = y_pos  + 20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_jobs_btn
		  first_job = TRUE
		  For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
			  If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
				  If first_job = TRUE Then
					  Text 40, y_pos, 20, 10, "Name"
					  Text 160, y_pos, 55, 10, "Employer Name"
					  Text 270, y_pos, 35, 10, "Start Date"
					  Text 330, y_pos, 35, 10, "Seasonal"
					  Text 375, y_pos, 30, 10, "Amount"
					  Text 425, y_pos, 50, 10, "How often?"
					  y_pos = y_pos  + 10
					  first_job = FALSE
				  End If
				  ComboBox 40, y_pos, 110, 45, all_the_clients, NEW_EARNED_ARRAY(earned_client, each_job)
				  EditBox 155, y_pos, 105, 15, NEW_EARNED_ARRAY(earned_source, each_job)
				  EditBox 270, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_start_date, each_job)
				  DropListBox 330, y_pos, 40, 45, " "+chr(9)+"No"+chr(9)+"Yes", NEW_EARNED_ARRAY(earned_seasonal, each_job)
				  EditBox 375, y_pos, 45, 15, NEW_EARNED_ARRAY(earned_amount, each_job)
				  DropListBox 425, y_pos, 60, 45, "Select One..."+chr(9)+"4 - Weekly"+chr(9)+"3 - Biweekly"+chr(9)+"2 - Semi Monthly"+chr(9)+"1 - Monthly", NEW_EARNED_ARRAY(earned_freq, each_job)
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_INCOME_ARRAY(new_checkbox, each_job)
				  ' CheckBox 530, y_pos, 40, 10, "Detail", ALL_INCOME_ARRAY(update_checkbox, each_job)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_job = TRUE Then
			  Text 35, y_pos, 400, 10, "CSR form - no JOBS information has been added."
			  y_pos = y_pos + 20
		  Else
			y_pos = y_pos + 10
		  End If

		  GroupBox 15, y_pos + 5, 585, q_7_grp_len, "Q7. Does anyone get money or does anyone expect to get money from sources other than work?"
		  ' DropListBox 335, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_other_income
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q7 Deailts left Blank", q_7_details_blank_checkbox
		  DropListBox 335, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_other_income
		  y_pos = y_pos +20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_unea_btn
		  first_unea = TRUE

		  For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			  If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
				  If first_unea = TRUE Then
					  Text 30, y_pos, 25, 10, "Name"
					  Text 165, y_pos, 55, 10, "Type of Income"
					  Text 280, y_pos, 35, 10, "Start Date"
					  Text 335, y_pos, 35, 10, "Amount"
					  Text 390, y_pos, 55, 10, "How often recvd"
					  y_pos = y_pos + 10
					  first_unea = FALSE
				  End If
				  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_UNEARNED_ARRAY(unearned_client, each_unea)
				  ComboBox 165, y_pos, 110, 45, unea_type_list, NEW_UNEARNED_ARRAY(unearned_source, each_unea)   'unea_type
				  EditBox 280, y_pos, 50, 15, NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)    'unea_start_date
				  EditBox 335, y_pos, 50, 15, NEW_UNEARNED_ARRAY(unearned_amount, each_unea)    'unea_amount
				  DropListBox 390, y_pos, 90, 45, "Select One..."+chr(9)+"4 - Weekly"+chr(9)+"3 - Biweekly"+chr(9)+"2 - Semi Monthly"+chr(9)+"1 - Monthly", NEW_UNEARNED_ARRAY(unearned_freq, each_unea) 'unea_frequency
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_INCOME_ARRAY(new_checkbox, each_unea)
				  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_INCOME_ARRAY(update_checkbox, each_unea)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_unea = TRUE Then
			  Text 35, y_pos, 400, 10, "CSR form - no UNEA information has been added."
			  y_pos = y_pos + 20
		  Else
			y_pos = y_pos + 10
		  End If

		  ButtonGroup ButtonPressed
			PushButton 20, y_pos + 2, 200, 13, "Why do I have to answer these if in is not HC?", why_answer_btn
			PushButton 475, y_pos, 80, 15, "Go to Q9 - Q12", next_page_ma_btn
			CancelButton 555, y_pos, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation

		err_msg = "LOOP"

		If ButtonPressed = -1 Then ButtonPressed = next_page_ma_btn

		If all_questions_4_7_blank = "Questions 4 - 7 are completely blank." Then
			apply_for_ma = "Did not answer and details blank"
			ma_self_employed = "Did not answer and details blank"
			ma_start_working = "Did not answer and details blank"
			ma_other_income = "Did not answer and details blank"

			' q_4_details_blank_checkbox = checked
			' q_5_details_blank_checkbox = checked
			' q_6_details_blank_checkbox = checked
			' q_7_details_blank_checkbox = checked
		End If

		If ButtonPressed = add_memb_btn Then
			If NEW_MA_REQUEST_ARRAY(ma_request_client, 0) = "Select or Type" Then
				NEW_MA_REQUEST_ARRAY(ma_request_client, 0) = "Enter or Select Member"
			Else
				new_item = UBound(NEW_MA_REQUEST_ARRAY, 2) + 1
				ReDim Preserve NEW_MA_REQUEST_ARRAY(ma_request_notes, new_item)
				NEW_MA_REQUEST_ARRAY(ma_request_client, new_item) = "Enter or Select Member"
			End If
		End If
		If ButtonPressed = add_busi_btn Then
			ReDim Preserve NEW_EARNED_ARRAY(earned_notes, new_earned_counter)
			NEW_EARNED_ARRAY(earned_type, new_earned_counter) = "BUSI"
			NEW_EARNED_ARRAY(earned_prog_list, new_earned_counter) = "MA"
			new_earned_counter = new_earned_counter + 1
		End If
		If ButtonPressed = add_jobs_btn Then
			ReDim Preserve NEW_EARNED_ARRAY(earned_notes, new_earned_counter)
			NEW_EARNED_ARRAY(earned_type, new_earned_counter) = "JOBS"
			NEW_EARNED_ARRAY(earned_prog_list, new_earned_counter) = "MA"
			new_earned_counter = new_earned_counter + 1
		End If
		If ButtonPressed = add_unea_btn Then
			new_item = UBound(ALL_INCOME_ARRAY, 2) + 1
			ReDim Preserve NEW_UNEARNED_ARRAY(unearned_notes, new_unearned_counter)
			NEW_UNEARNED_ARRAY(unearned_type, new_unearned_counter) = "UNEA"
			NEW_UNEARNED_ARRAY(unearned_prog_list, new_unearned_counter) = "MA"
			new_unearned_counter = new_unearned_counter + 1
		End If
		If ButtonPressed = why_answer_btn Then
			explain_text = "This case may not have MA, MSP, or any HC active and you may have indicated that it is only for a SNAP Review, HOWEVER" & vbCr & vbCr
			explain_text = explain_text & "The form that was sent to the client STILL has these questions listed on it." & vbCr
			explain_text = explain_text & "We need to be looking at all information that the client reported, anything entered here may impact the benefits because it is now 'known to the agency'." & vbCr & vbCr
			explain_text = explain_text & "Though the client is not required to answer these questions, we are still required to review the entire form."
			' explain_text = explain_text & ""
			why_answer_when_not_HC_msg = MsgBOx(explain_text, vbInformation + vbOKonly, "No HC on the case")
		End If

		If ButtonPressed = next_page_ma_btn Then
			questions_answered = TRUE
			err_msg = ""

			If apply_for_ma = "Select One..." Then questions_answered = FALSE
			If ma_self_employed = "Select One..." Then questions_answered = FALSE
			If ma_start_working = "Select One..." Then questions_answered = FALSE
			If ma_other_income = "Select One..." Then questions_answered = FALSE

			If questions_answered = FALSE Then
				err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form."
				If apply_for_ma = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 4 about applying for someone not currently getting MA coverage."
				If ma_self_employed = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 5 about anyone being self-employed."
				If ma_start_working = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 6 about anyone working."
				If ma_other_income = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Question 7 about unearned income."
			End If

			q_4_details_entered = FALSE
			q_5_details_entered = FALSE
			q_6_details_entered = FALSE
			q_7_details_entered = FALSE
			' If q_4_details_blank_checkbox = unchecked Then
			If InStr(apply_for_ma, "details listed below") <> 0 Then
				For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
					If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Select or Type" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" and NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "Enter or Select Member" Then
						q_4_details_entered = TRUE
					End If
				Next
				' If q_4_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No details of a person requesting MA for someone not getting coverage now (Question 4). Either enter information about which members are requesting MA coverage or check the box to indicate this portion of the form was left blank."
				' If q_4_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 4 - Someone getting MA coverage. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_4_details_entered = TRUE
			End If
			' If q_5_details_blank_checkbox = unchecked Then
			If InStr(ma_self_employed, "details listed below") <> 0 Then
				For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
					If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
						q_5_details_entered = TRUE
					End If
				Next
				' If q_5_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No Self Employment information has been entered (Question 5). Either enter BUSI details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_5_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 5 - Self-Employment Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_5_details_entered = TRUE
			End If
			' If q_6_details_blank_checkbox = unchecked Then
			If InStr(ma_start_working, "details listed below") <> 0 Then
				For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
					If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
						q_6_details_entered = TRUE
					End If
				Next
				' If q_6_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Job information has been entered (Question 6). Either enter JOBS details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_6_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 6 - Job Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_6_details_entered = TRUE
			End If
			' If q_7_details_blank_checkbox = unchecked Then
			If InStr(ma_other_income, "details listed below") <> 0 Then
				For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
					If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
						q_7_details_entered = TRUE
					End If
				Next
				' If q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 7). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 7 - Unearned Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_7_details_entered = TRUE
			End If

			If q_4_details_entered = FALSE OR q_5_details_entered = FALSE OR q_6_details_entered = FALSE  OR q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "The folowing questions need review. The answer indicates there was detail provided but no detail was listed for:"
			If q_4_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 4 - Someone getting MA coverage."
			If q_5_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 5 - Self-Employment Income."
			If q_6_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 6 - Job Income. "
			If q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 7 - Unearned Income."
			If q_4_details_entered = FALSE OR q_5_details_entered = FALSE OR q_6_details_entered = FALSE  OR q_7_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-- Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form. --"

			If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
			If err_msg = "" Then csr_dlg_q_4_7_cleared = TRUE
		End If
	Loop until err_msg = ""
	show_csr_dlg_q_4_7 = FALSE

end function

function csr_dlg_q_9_12()
	Do
		dlg_len = 205
		q_9_grp_len = 30
		q_10_grp_len = 30
		q_11_grp_len = 30
		q_12_grp_len = 30
		For new_assets_listed = 0 to UBound(NEW_ASSET_ARRAY, 2)
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "CASH" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_9_grp_len = q_9_grp_len + 20
				' MsgBox ALL_ASSETS_ARRAY(category_const, assets_on_case)
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "ACCT" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_9_grp_len = q_9_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_10_grp_len = q_10_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_11_grp_len = q_11_grp_len + 20
			End If
			If NEW_ASSET_ARRAY(asset_type, new_assets_listed) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, new_assets_listed) = "MA" Then
				dlg_len = dlg_len + 20
				q_12_grp_len = q_12_grp_len + 20
			End If
		Next
		y_pos = 25
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 610, dlg_len, "MA CSR Asset Questions"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 9 - 12:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 9 - 12 are completely blank.", all_questions_9_12_blank

		  GroupBox 15, y_pos + 5, 585, q_9_grp_len, "Q9. Does anyone have cash, a savings or checking account, or a certificate of deposit?"
		  ' DropListBox 330, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_liquid_assets
		  ' CheckBox 430, y_pos + 5, 75, 10, "Q9 Deailts left Blank", q_9_details_blank_checkbox
		  DropListBox 330, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_liquid_assets
		  y_pos = y_pos +20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_acct_btn
		  first_account = TRUE

		  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
			If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			  If first_account = TRUE Then
				  Text 30, y_pos, 55, 10, "Owner(s) Name"
				  Text 165, y_pos, 25, 10, "Type"
				  Text 285, y_pos, 50, 10, "Bank Name"
				  y_pos = y_pos + 10
				  first_account = FALSE
			  End If
			  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset) 'liquid_asset_member'
			  ComboBox 165, y_pos, 115, 40, account_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)'liquid_asst_type
			  EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)'liquid_asset_name
			  ' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)    'new_checkbox
			  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)'update_checkbox
			  y_pos = y_pos + 20
			End If
		  Next
		  If first_account = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no ACCT information has been added."
			y_pos = y_pos + 10
		  End If

		  y_pos = y_pos +10
		  GroupBox 15, y_pos + 5, 585, q_10_grp_len, "Q10. Does anyone own or co-own securities or other assets?"
		  ' DropListBox 295, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_security_assets
		  ' CheckBox 430, y_pos + 5, 80, 10, "Q10 Deailts left Blank", q_10_details_blank_checkbox
		  DropListBox 295, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_security_assets
		  y_pos = y_pos +  20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_secu_btn

		  first_secu = TRUE
		  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
			If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
				If first_secu = TRUE Then
					Text 30, y_pos, 55, 10, "Owner(s) Name"
					Text 165, y_pos, 25, 10, "Type"
					Text 285, y_pos, 50, 10, "Bank Name"
					y_pos = y_pos + 10
					first_secu = FALSE
				End If
				ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset) 'security_asset_member
				ComboBox 165, y_pos, 115, 40, security_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'security_asset_type
				EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)   'security_asset_name
				' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)
				' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)
				y_pos = y_pos + 20
			End If
		  Next
		  If first_secu = TRUE Then
			  Text 30, y_pos, 250, 10, "CSR form - no SECU information has been added."
			  y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  GroupBox 15, y_pos + 5, 585, q_11_grp_len, "Q11. Does anyone own a vehicle?"
		  ' DropListBox 250, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_vehicle
		  ' CheckBox 430, y_pos + 5, 80, 10, "Q11 Deailts left Blank", q_11_details_blank_checkbox
		  DropListBox 250, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_vehicle
		  y_pos = y_pos + 20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_cars_btn
		  first_car = TRUE
		  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
			  If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
				  If first_car = TRUE Then
					  Text 30, y_pos, 55, 10, "Owner(s) Name"
					  Text 165, y_pos, 25, 10, "Type"
					  Text 285, y_pos, 75, 10, "Year/Make/Model"
					  y_pos = y_pos + 10
					  first_car = FALSE
				  End If
				  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset)     'vehicle_asset_member
				  ComboBox 165, y_pos, 115, 40, cars_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'vehicle_asset_type
				  EditBox 285, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_year_make_model, each_asset)  'vehicle_asset_name
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)
				  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_car = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no CARS information has been added."
			y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  GroupBox 15, y_pos + 5, 585, q_12_grp_len, "Q12. Does anyone own or co-own any real estate?"
		  ' DropListBox 280, y_pos, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_real_assets
		  ' CheckBox 430, y_pos + 5, 80, 10, "Q12 Deailts left Blank", q_12_details_blank_checkbox
		  DropListBox 280, y_pos, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_real_assets
		  y_pos = y_pos + 20
		  ButtonGroup ButtonPressed
			PushButton 540, y_pos - 15, 50, 10, "Add Another", add_rest_btn
		  first_home = TRUE
		  For each_asset = 0 to Ubound(NEW_ASSET_ARRAY, 2)
			  If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
				  If first_home = TRUE Then
					  Text 30, y_pos, 55, 10, "Owner(s) Name"
					  Text 165, y_pos, 25, 10, "Address"
					  Text 320, y_pos, 75, 10, "Type of Property"
					  y_pos = y_pos + 10
					  first_home = FALSE
				  End If
				  ComboBox 30, y_pos, 130, 45, all_the_clients, NEW_ASSET_ARRAY(asset_client, each_asset)     'property_asset_member
				  EditBox 165, y_pos, 150, 15, NEW_ASSET_ARRAY(asset_address, each_asset)      'property_asset_address
				  ComboBox 320, y_pos, 150, 40, rest_list, NEW_ASSET_ARRAY(asset_acct_type, each_asset)     'property_asset_type
				  ' CheckBox 495, y_pos, 30, 10, "New", ALL_ASSETS_ARRAY(new_checkbox, each_asset)
				  ' CheckBox 530, y_pos, 55, 10, "Update/Detail", ALL_ASSETS_ARRAY(update_checkbox, each_asset)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_home = TRUE Then
			Text 30, y_pos, 250, 10, "CSR form - no REST information has been added."
			y_pos = y_pos + 10
		  End If
		  y_pos = y_pos + 10

		  ButtonGroup ButtonPressed
			PushButton 415, y_pos, 80, 15, "Go Back to Q4 - Q7", back_to_ma_dlg_1
			PushButton 495, y_pos, 60, 15, "Continue", continue_btn
			CancelButton 555, y_pos, 50, 15
		EndDialog

		err_msg = "LOOP"

		dialog Dialog1
		cancel_confirmation

		' MsgBox ButtonPressed & " - 1 - "
		If ButtonPressed = -1 Then ButtonPressed = continue_btn

		If ButtonPressed = add_acct_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "ACCT"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If
		If ButtonPressed = add_secu_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "SECU"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If
		If ButtonPressed = add_cars_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "CARS"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If
		If ButtonPressed = add_rest_btn Then
			ReDim Preserve NEW_ASSET_ARRAY(asset_notes, new_asset_counter)
			NEW_ASSET_ARRAY(asset_type, new_asset_counter) = "REST"
			NEW_ASSET_ARRAY(asset_prog_list, new_asset_counter) = "MA"
			new_asset_counter = new_asset_counter + 1
		End If

		If all_questions_9_12_blank = "Questions 9 - 12 are completely blank." Then
			ma_liquid_assets = "Did not answer and details blank"
			ma_security_assets = "Did not answer and details blank"
			ma_vehicle = "Did not answer and details blank"
			ma_real_assets = "Did not answer and details blank"

			' q_9_details_blank_checkbox = checked
			' q_10_details_blank_checkbox = checked
			' q_11_details_blank_checkbox = checked
			' q_12_details_blank_checkbox = checked
		End If

		If ButtonPressed = continue_btn Then
			questions_answered = TRUE
			err_msg = ""

			If ma_liquid_assets = "Select One..." Then questions_answered = FALSE
			If ma_security_assets = "Select One..." Then questions_answered = FALSE
			If ma_vehicle = "Select One..." Then questions_answered = FALSE
			If ma_real_assets = "Select One..." Then questions_answered = FALSE

			If questions_answered = FALSE Then
				err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form."

				If ma_liquid_assets = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
				If ma_security_assets = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
				If ma_vehicle = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
				If ma_real_assets = "Select One..." Then err_msg = err_msg & vbNewLine & "   - "
			End If

			q_9_details_entered = FALSE
			q_10_details_entered = FALSE
			q_11_details_entered = FALSE
			q_12_details_entered = FALSE
			' If q_9_details_blank_checkbox = unchecked Then
			If InStr(ma_liquid_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" AND NEW_ASSET_ARRAY(ma_request_client, each_asset) <> "Enter or Select Member" Then
						q_9_details_entered = TRUE
					End If
				Next
				' If q_9_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No details of a person requesting MA for someone not getting coverage now (Question 9). Either enter information about which members are requesting MA coverage or check the box to indicate this portion of the form was left blank."
				' If q_9_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 9 - Liquid Assets. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_9_details_entered = TRUE
			End If
			' If q_10_details_blank_checkbox = unchecked Then
			If InStr(ma_security_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_10_details_entered = TRUE
					End If
				Next
				' If q_10_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* No Self Employment information has been entered (Question 10). Either enter BUSI details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_10_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 10 - Securities. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_10_details_entered = TRUE
			End If
			' If q_11_details_blank_checkbox = unchecked Then
			If InStr(ma_vehicle, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_11_details_entered = TRUE
					End If
				Next
				' If q_11_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Job information has been entered (Question 11). Either enter JOBS details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_11_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 11 - Vehicles. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_11_details_entered = TRUE
			End If
			' If q_12_details_blank_checkbox = unchecked Then
			If InStr(ma_real_assets, "details listed below") <> 0 Then
				For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
					If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
						q_12_details_entered = TRUE
					End If
				Next
				' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
				' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 12 - Real Estate. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
			Else
				q_12_details_entered = TRUE
			End If

			If q_9_details_entered = FALSE OR q_10_details_entered = FALSE OR q_11_details_entered = FALSE  OR q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "The folowing questions need review. The answer indicates there was detail provided but no detail was listed for:"
			If q_9_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 9 - Liquid Assets."
			If q_10_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-  Question 10 - Securities."
			If q_11_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 11 - Vehicles."
			If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 12 - Real Estate."
			If q_9_details_entered = FALSE OR q_10_details_entered = FALSE OR q_11_details_entered = FALSE  OR q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-- Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form. --"


			If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue:" & vbNewLine & err_msg
			If err_msg = "" Then 	csr_dlg_q_9_12_cleared = TRUE
		End If

		If ButtonPressed = back_to_ma_dlg_1 Then
			' MsgBox ButtonPressed & " - 2 - "
			show_csr_dlg_q_4_7 = TRUE
			show_csr_dlg_q_13 = FALSE
			show_csr_dlg_q_15_19 = FALSE
			show_csr_dlg_sig = FALSE
			show_confirmation = FALSE
			err_msg = ""
		End If
	Loop until err_msg = ""
	show_csr_dlg_q_9_12 = FALSE

end function

function csr_dlg_q_13()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 610, 80, "MA CSR Changes"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 13:"

		  Text 25, 25, 135, 10, "Q13. Do you have any changes to report?"
		  ' DropListBox 160, 20, 75, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes"+chr(9)+"Did not answer", ma_other_changes
		  ' CheckBox 30, 60, 300, 10, "Check here if client left the changes to report field on the form blank.", changes_reported_blank_checkbox
		  DropListBox 160, 20, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", ma_other_changes
		  EditBox 30, 40, 555, 15, other_changes_reported

		  ButtonGroup ButtonPressed
			PushButton 255, 60, 100, 15, "Back to Q 4-7", back_to_ma_dlg_1
			PushButton 355, 60, 100, 15, "Back to Q 9 - 12", back_to_ma_dlg_2
			PushButton 455, 60, 100, 15, "Finish MA Questions", finish_ma_questions
			CancelButton 555, 60, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation
		If ButtonPressed = -1 Then ButtonPressed = finish_ma_questions

		If ButtonPressed = back_to_ma_dlg_1 Then
			show_csr_dlg_q_4_7 = TRUE
			show_csr_dlg_q_15_19 = FALSE
			show_csr_dlg_sig = FALSE
			show_confirmation = FALSE
		End If
		If ButtonPressed = back_to_ma_dlg_2 Then
			show_csr_dlg_q_9_12 = TRUE
			show_csr_dlg_q_15_19 = FALSE
			show_csr_dlg_sig = FALSE
			show_confirmation = FALSE
		End If
		If ButtonPressed = finish_ma_questions Then
			show_ma_dlg_three = FALSE

			questions_answered = TRUE

			If trim(other_changes_reported) <> "" Then details_shown = TRUE
			If Instr(ma_other_changes, "details blank") <> 0 Then details_shown = TRUE

			If ma_other_changes = "Select One..." Then questions_answered = FALSE

			If questions_answered = FALSE Then
				err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form."

				If ma_other_changes = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Indicate what the client entered for Question 13."
			Else
				If details_shown = FALSE Then err_msg = err_msg & vbNewLine & "* You must either enter what the client wrote in for Question 13 or check the box to indicate if if was blank."
			End If
			If trim(other_changes_reported) <> "" AND changes_reported_blank_checkbox = checked Then err_msg = err_msg & vbNewLine & "* You entered detail in what the client wrote and indicated it was blank using the checkbox, please update as only one of these should be completed."

			If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
			If err_msg = "" Then csr_dlg_q_13_cleared = TRUE
		End If
	Loop until err_msg = ""
	show_csr_dlg_q_13 = FALSE
	' MsgBox "Q 13 Cleared - " & csr_dlg_q_13_cleared
end function

function csr_dlg_q_15_19()
	Do
		err_msg = ""

		dlg_len = 190
		q_15_grp_len = 30
		q_16_grp_len = 25
		q_17_grp_len = 25
		q_18_grp_len = 25

		For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
			If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
				dlg_len = dlg_len + 20
				q_16_grp_len = q_16_grp_len + 20
			End If
		Next

		For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
				dlg_len = dlg_len + 20
				q_17_grp_len = q_17_grp_len + 20
			End If
		Next

		For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
			If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
				dlg_len = dlg_len + 20
				q_18_grp_len = q_18_grp_len + 20
			End If
		Next
		' dlg_len = dlg_len + UBound(NEW_EARNED_ARRAY, 2) * 20
		' dlg_len = dlg_len + UBound(NEW_UNEARNED_ARRAY, 2) * 20
		' dlg_len = dlg_len + UBound(NEW_CHILD_SUPPORT_ARRAY, 2) * 20
		' q_15_grp_len = 50
		' q_16_grp_len = 45 + UBound(NEW_EARNED_ARRAY, 2) * 20
		' q_17_grp_len = 45 + UBound(NEW_UNEARNED_ARRAY, 2) * 20
		' q_18_grp_len = 45 + UBound(NEW_CHILD_SUPPORT_ARRAY, 2) * 20

		y_pos = 95

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 615, dlg_len, "SNAP CSR Question Details"
		  Text 10, 10, 180, 10, "Enter the answers from the CSR form, questions 15 - 19:"
		  Text 195, 10, 210, 10, "If all questions and details have been left blank, indicate that here:"
		  DropListBox 405, 5, 200, 15, "Enter Question specific information below"+chr(9)+"Questions 15 - 19 are completely blank.", all_questions_15_19_blank
		  GroupBox 10, 30, 600, q_15_grp_len, "Q15. Has your household moved since your last application or in the past six months?"
		  ' DropListBox 305, 25, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_fifteen_form_answer
		  DropListBox 305, 25, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", quest_fifteen_form_answer
		  Text 25, 45, 105, 10, "New Rent or Mortgage Amount:"
		  EditBox 130, 40, 65, 15, new_rent_or_mortgage_amount
		  CheckBox 220, 45, 50, 10, "Heat/AC", heat_ac_checkbox
		  CheckBox 275, 45, 50, 10, "Electricity", electricity_checkbox
		  CheckBox 345, 45, 50, 10, "Telephone", telephone_checkbox
		  Text 400, 45, 80, 10, "Did client attach proof?"
		  DropListBox 480, 40, 125, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", shel_proof_provided
		  GroupBox 10, 70, 490, q_16_grp_len, "Q16 Has there been a change in EARNED INCOME?"
		  ' DropListBox 190, 65, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_sixteen_form_answer
		  ' CheckBox 310, 70, 85, 10, "Q16 Deailts left Blank", q_16_details_blank_checkbox
		  DropListBox 190, 65, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", quest_sixteen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 440, 70, 50, 10, "Add Another", add_snap_earned_income_btn
		  first_earned = TRUE
		  For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
			  If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
				  If first_earned = TRUE Then
					  Text 15, 85, 20, 10, "Client"
					  Text 130, 85, 100, 10, "Employer (or Business Name)"
					  Text 265, 85, 50, 10, "Change Date"
					  Text 320, 85, 35, 10, "Amount"
					  Text 375, 85, 40, 10, "Frequency"
					  Text 445, 85, 25, 10, "Hours"
					  first_earned = FALSE
				  End If
				  ComboBox 15, y_pos, 110, 45, all_the_clients, NEW_EARNED_ARRAY(earned_client, the_earned)
				  EditBox 130, y_pos, 130, 15, NEW_EARNED_ARRAY(earned_source, the_earned)
				  EditBox 265, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_change_date, the_earned)
				  EditBox 320, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_amount, the_earned)
				  DropListBox 375, y_pos, 65, 45, "Select One..."+chr(9)+"Weekly"+chr(9)+"BiWeekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", NEW_EARNED_ARRAY(earned_freq, the_earned)
				  EditBox 445, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_hours, the_earned)
				  y_pos = y_pos + 20
			  End If
		  Next
		  y_pos = y_pos + 10
		  GroupBox 10, y_pos, 490, q_17_grp_len, "Q17. Has there been a change in UNEARNED INCOME?"
		  ' DropListBox 205, y_pos - 5, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_seventeen_form_answer
		  ' CheckBox 310, y_pos, 85, 10, "Q17 Deailts left Blank", q_17_details_blank_checkbox
		  DropListBox 205, y_pos - 5, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", quest_seventeen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 440, y_pos, 50, 10, "Add Another", add_snap_unearned_btn
		  y_pos = y_pos + 15
		  first_unearned = TRUE
		  For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
			  If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
				  If first_unearned = TRUE Then
					  Text 15, y_pos, 20, 10, "Client"
					  Text 145, y_pos, 100, 10, "Type and Source"
					  Text 280, y_pos, 50, 10, "Change Date"
					  Text 340, y_pos, 35, 10, "Amount"
					  Text 405, y_pos, 40, 10, "Frequency"
					  y_pos = y_pos + 10
					  first_unearned = FALSE
				  End If
				  ComboBox 15, y_pos, 125, 45, all_the_clients, NEW_UNEARNED_ARRAY(unearned_client, the_unearned)
				  EditBox 145, y_pos, 130, 15, NEW_UNEARNED_ARRAY(unearned_source, the_unearned)
				  EditBox 280, y_pos, 55, 15, NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned)
				  EditBox 340, y_pos, 60, 15, NEW_UNEARNED_ARRAY(unearned_amount, the_unearned)
				  DropListBox 405, y_pos, 90, 45, "Select One..."+chr(9)+"Weekly"+chr(9)+"BiWeekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", NEW_UNEARNED_ARRAY(unearned_freq, the_unearned)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_unearned = TRUE Then y_pos = y_pos + 10
		  y_pos = y_pos + 10
		  GroupBox 10, y_pos, 490, q_18_grp_len, "Q18 Has there been a change in CHILD SUPPORT?"
		  ' DropListBox 190, y_pos - 5, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_eighteen_form_answer
		  ' CheckBox 310, y_pos, 85, 10, "Q18 Deailts left Blank", q_18_details_blank_checkbox
		  DropListBox 190, y_pos - 5, 150, 45, "Select One..."+chr(9)+"No and details blank"+chr(9)+"No but details listed below"+chr(9)+"Yes but details blank"+chr(9)+"Yes and details listed below"+chr(9)+"Did not answer and details blank"+chr(9)+"Did not answer but details listed below", quest_eighteen_form_answer
		  ButtonGroup ButtonPressed
			PushButton 440, y_pos, 50, 10, "Add Another", add_snap_cs_btn
		  y_pos = y_pos + 15

		  first_cs = TRUE
		  For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
			  If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
				  If first_cs = TRUE Then
					  Text 15, y_pos, 85, 10, "Name of person paying"
					  Text 220, y_pos, 35, 10, "Amount"
					  Text 295, y_pos, 65, 10, "Currently Paying?"
					  y_pos = y_pos + 10
					  first_cs = FALSE
				  End If
				  EditBox 15, y_pos, 200, 15, NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs)
				  EditBox 220, y_pos, 65, 15, NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs)
				  DropListBox 295, y_pos, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs)
				  y_pos = y_pos + 20
			  End If
		  Next
		  If first_cs = TRUE Then y_pos = y_pos + 10
		  y_pos = y_pos + 10
		  Text 10, y_pos, 345, 10, "Q19. Did you work 20 hours each week, for an average of 80 hours per month during the past six months?"
		  DropListBox 355, y_pos - 5, 100, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Did not answer", quest_nineteen_form_answer
		  ' y_pos = y_pos + 15
		  ButtonGroup ButtonPressed
			PushButton 505, y_pos-5, 50, 15, "Continue", continue_btn
			CancelButton 555, y_pos-5, 50, 15
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If all_questions_15_19_blank = "Questions 15 - 19 are completely blank." Then
			quest_fifteen_form_answer = "Did not answer and details blank"
			quest_sixteen_form_answer = "Did not answer and details blank"
			quest_seventeen_form_answer = "Did not answer and details blank"
			quest_eighteen_form_answer = "Did not answer and details blank"
			quest_nineteen_form_answer = "Did not answer"

			' q_16_details_blank_checkbox = checked
			' q_17_details_blank_checkbox = checked
			' q_18_details_blank_checkbox = checked
		End If

		If quest_fifteen_form_answer = "Select One..." OR quest_sixteen_form_answer = "Select One..." OR quest_seventeen_form_answer = "Select One..." OR quest_eighteen_form_answer = "Select One..." OR quest_nineteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & "* All of the questions must be answered with the answers from the CSR Form." & vbNewLine
		If quest_fifteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 15 (Has the household moved?)."
		If quest_sixteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 16 (Has anyone had a change in Earned income?)."
		If quest_seventeen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 17 (Has anyone had a change in Unearned income?)."
		If quest_eighteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 18 (Has there been a change in Child Support income?)."
		If quest_nineteen_form_answer = "Select One..." Then err_msg = err_msg & vbNewLine & " - Indicate the answer on the CSR form for Question 19 (Have you worked 80 hours per month?)."

		q_15_details_entered = FALSE
		q_16_details_entered = FALSE
		q_17_details_entered = FALSE
		q_18_details_entered = FALSE

		If InStr(quest_fifteen_form_answer, "details listed below") <> 0 Then

			new_rent_or_mortgage_amount = trim(new_rent_or_mortgage_amount)
			If new_rent_or_mortgage_amount <> "" Then q_15_details_entered = TRUE
			If heat_ac_checkbox = CHECKED Then q_15_details_entered = TRUE
			If electricity_checkbox = CHECKED Then q_15_details_entered = TRUE
			If telephone_checkbox = CHECKED Then q_15_details_entered = TRUE

			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_15_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 15 - Shelter and Utilities Expenses. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_15_details_entered = TRUE
		End If

		If InStr(quest_sixteen_form_answer, "details listed below") <> 0 Then
			For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
				If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
					q_16_details_entered = TRUE
				End If
			Next
			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_16_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 16 - Earned Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_16_details_entered = TRUE
		End If

		If InStr(quest_seventeen_form_answer, "details listed below") <> 0 Then
			For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
				If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
					q_17_details_entered = TRUE
				End If
			Next
			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_17_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 17 - Unearned Income. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_17_details_entered = TRUE
		End If

		If InStr(quest_eighteen_form_answer, "details listed below") <> 0 Then
			For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
				If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
					q_18_details_entered = TRUE
				End If
			Next
			' If q_12_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* * No Unearned Income information has been entered (Question 12). Either enter UNEA details from the CSR Form or check the box to indicate this portion of the form was left blank."
			' If q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "* Question 18 - Child Support. The answer indicates there was detail provided but no detail was listed. Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form."
		Else
			q_18_details_entered = TRUE
		End If

		If q_15_details_entered = FALSE OR q_16_details_entered = FALSE OR q_17_details_entered = FALSE  OR q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "The folowing questions need review. The answer indicates there was detail provided but no detail was listed for:"
		If q_15_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 15 - Shelter and Utilities Expenses."
		If q_16_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 16 - Earned Income."
		If q_17_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 17 - Unearned Income."
		If q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "- Question 18 - Child Support."
		If q_15_details_entered = FALSE OR q_16_details_entered = FALSE OR q_17_details_entered = FALSE  OR q_18_details_entered = FALSE Then err_msg = err_msg & vbNewLine & "-- Update the droplist answer to indicate there are no details or add details about who is requesting MA coverage from the form. --"


		If ButtonPressed = add_snap_earned_income_btn Then
			ReDim Preserve NEW_EARNED_ARRAY(earned_notes, new_earned_counter)
			NEW_EARNED_ARRAY(earned_prog_list, new_earned_counter) = "SNAP"
			new_earned_counter = new_earned_counter + 1
			err_msg = "LOOP" & err_msg
		End If

		If ButtonPressed = add_snap_unearned_btn Then
			new_item = UBound(NEW_UNEARNED_ARRAY, 2) + 1
			ReDim Preserve NEW_UNEARNED_ARRAY(unearned_notes, new_unearned_counter)
			NEW_UNEARNED_ARRAY(unearned_prog_list, new_unearned_counter) = "SNAP"
			new_unearned_counter = new_unearned_counter + 1
			err_msg = "LOOP" & err_msg
		End If

		If ButtonPressed = add_snap_cs_btn Then
			If NEW_CHILD_SUPPORT_ARRAY(cs_current, 0) = "" THen
				NEW_CHILD_SUPPORT_ARRAY(cs_current, 0) = "Select..."
			Else
				new_item = UBound(NEW_CHILD_SUPPORT_ARRAY, 2) + 1
				ReDim Preserve NEW_CHILD_SUPPORT_ARRAY(cs_notes, new_item)
			End If
			err_msg = "LOOP" & err_msg

		End If

		If ButtonPressed = -1 Then ButtonPressed = continue_btn

		If err_msg <> "" AND left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
		' MsgBox show_two & vbNewLine & "line 1480"
		' Loop until leave_ma_questions = TRUE
	Loop until err_msg = ""
	show_csr_dlg_q_15_19 = FALSE
	csr_dlg_q_15_19_cleared = TRUE
end function

function csr_dlg_sig()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 201, 140, "Form dates and signatures"
		  EditBox 135, 30, 60, 15, csr_form_date
		  DropListBox 135, 55, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_yn
		  DropListBox 135, 75, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_dated_yn
		  ButtonGroup ButtonPressed
		    PushButton 35, 120, 105, 15, "Complete CSR Form Detail", complete_csr_questions
		    CancelButton 145, 120, 50, 15
		  Text 60, 35, 70, 10, "Date form Received:"
		  Text 10, 60, 120, 10, "Has the client signed the CSR Form?"
		  Text 10, 80, 120, 10, "Has the client dated the CSR Form?"
		  Text 10, 10, 115, 15, "Answer if the CSR Form has been signed and dated by the client:"
		  Text 5, 95, 190, 20, "Note: During the COVID Peacetime emergency verbal signatures over the phone count as 'Yes' for signed/dated."
		EndDialog

		dialog Dialog1

		cancel_confirmation

		If IsDate(csr_form_date) = FALSE Then
			err_msg = err_msg & vbNewLine & "* Enter a valid date for the date the form was received."
		Else
			If DateDiff("d", date, csr_form_date) > 0 Then err_msg = err_msg & vbNewLine & "* The date of the CSR form is listed as a future date, a form cannot be listed as received inthe future, please review the form date."
		End If
		If client_signed_yn = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate if the client has signed the form correctly by selecting 'yes' or 'no'."
		If client_dated_yn = "Select..." Then err_msg = err_msg & vbNewLine & "* Indicate if the form has been dated correctly by selecting 'yes' or 'no'."

		If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg

	Loop until err_msg = ""
	show_csr_dlg_sig = FALSE
	csr_dlg_sig_cleared = TRUE
end function

function confirm_csr_form_dlg()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 751, 370, "CSR Form Information"
	  If show_buttons_on_confirmation_dlg = TRUE Then
		  DropListBox 255, 350, 190, 50, "Indicate the form information"+chr(9)+"NO - the information here is different"+chr(9)+"YES - This is the information on the CSR Form", confirm_csr_form_information
		  ButtonGroup ButtonPressed
		    OkButton 645, 350, 50, 15
		    CancelButton 695, 350, 50, 15
			PushButton 15, 327, 165, 13, "Fix Page One Information", back_to_dlg_addr
			PushButton 200, 327, 165, 13, "Fix Page Two Information", back_to_dlg_ma_income
			PushButton 385, 327, 165, 13, "Fix Page Three Information", back_to_dlg_ma_asset
			PushButton 570, 270, 165, 13, "Fix Page Four Information", back_to_dlg_snap
			PushButton 570, 327, 165, 13, "Fix Page Five Information", back_to_dlg_sig
	  Else
		  ButtonGroup ButtonPressed
			OkButton 695, 350, 50, 15
	  End If
	  GroupBox 5, 5, 185, 340, "Page 1"
	  Text 10, 20, 105, 10, "1. Name and Address"
	  Text 20, 35, 160, 10, "Name:" & client_on_csr_form
	  Text 20, 50, 70, 10, "Residence Address"
	  If new_resi_addr_entered = TRUE Then
		  Text 25, 65, 110, 10, new_resi_one
		  Text 25, 75, 110, 10, new_resi_city & ", " & new_resi_state & " " & new_resi_zip
		  y_pos_1 = 85
	  Else
		  Text 25, 65, 110, 10, resi_line_one
		  If resi_line_two = "" Then
			Text 25, 75, 110, 10, resi_city & ", " & resi_state & " " & resi_zip
			y_pos_1 = 85
		  Else
			Text 25, 75, 110, 10, resi_line_two
			Text 25, 85, 110, 10, resi_city & ", " & resi_state & " " & resi_zip
			y_pos_1 = 95
		  End If
	  End If
	  If residence_address_match_yn = "Yes - the addresses are the same." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - YES"
	  If residence_address_match_yn = "No - there is a difference." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If residence_address_match_yn = "No - New Address Entered" Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If residence_address_match_yn = "RESI Address not Provided" Then Text 100, y_pos_1, 75, 10, "BLANK RESI ADDR"

	  Text 20, y_pos_1 + 15, 70, 10, "Mailing Address"
	  y_pos_1 = y_pos_1 + 30
	  If new_mail_addr_entered = TRUE Then
		  Text 25, y_pos_1, 110, 10, new_mail_one
		  y_pos_1 = y_pos_1 + 10
		  Text 25, y_pos_1, 110, 10, new_mail_city & ", " & new_mail_state & " " & new_mail_zip
		  y_pos_1 = y_pos_1 + 10
	  Else
		  If mail_line_one = "" Then
			  Text 25, y_pos_1, 110, 10, "NO MAILING ADDRESS LISTED"
			  y_pos_1 = y_pos_1 + 15
		  Else
			  Text 25, y_pos_1, 110, 10, mail_line_one
			  y_pos_1 = y_pos_1 + 10
			  If mail_line_two = "" Then
				Text 25, y_pos_1, 110, 10, mail_city & ", " & mail_state & " " & mail_zip
				y_pos_1 = y_pos_1 + 10
			  Else
				Text 25, y_pos_1, 110, 10, mail_line_two
				Text 25, y_pos_1 + 10, 110, 10, mail_city & ", " & mail_state & " " & mail_zip
				y_pos_1 = y_pos_1 + 20
			  End If
		  End If
	  End If
	  If mailing_address_match_yn = "Yes - the addresses are the same." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - YES"
	  If mailing_address_match_yn = "No - there is a difference." Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If mailing_address_match_yn = "No - New Address Entered" Then Text 100, y_pos_1, 75, 10, "Matches MAXIS - NO"
	  If mailing_address_match_yn = "MAIL Address not Provided" Then Text 100, y_pos_1, 75, 10, "BLANK MAIL ADDR"

	  y_pos_1 = y_pos_1 + 15
	  Text 10, y_pos_1, 110, 20, "2. Has anyone moved in or out of your home in the past six months?"
	  ' Text 20, 160, 115, 10, "your home in the past six months?"
	  Text 150, y_pos_1, 35, 10, quest_two_move_in_out
	  y_pos_1 = y_pos_1 + 25
	  pers_list_count = 1

	  For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
		  If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked OR ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then
			  Text 20, y_pos_1, 50, 10, "Person 1" & pers_list_count
			  If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then Text 140, y_pos_1, 40, 10, "MOVED OUT"
			  If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then Text 140, y_pos_1, 40, 10, "MOVED IN"
			  Text 25, y_pos_1 + 10, 155, 10, "Name:" & ALL_CLIENTS_ARRAY(memb_first_name, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb)
			  If len(ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_memb)) > 4 Then Text 25, y_pos_1 + 20, 155, 10, "Relationship:" & right(ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_memb), len(ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_memb)) - 5)
			  ' Text 25, 205, 155, 10, "Date of Change:"
			  ' Text 25, 215, 155, 10, "other"
			  pers_list_count = pers_list_count + 1
			  y_pos_1 = y_pos_1 + 40
		  End If
	  Next
	  For new_memb_counter = 0 to UBOUND(NEW_MEMBERS_ARRAY, 2)
		  If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter) = checked OR NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter) = checked Then
			  Text 20, y_pos_1, 50, 10, "Person 1" & pers_list_count
			  If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter) = checked Then Text 140, y_pos_1, 40, 10, "MOVED OUT"
			  If NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter) = checked Then Text 140, y_pos_1, 40, 10, "MOVED IN"
			  Text 25, y_pos_1 + 10, 155, 10, "Name:" & NEW_MEMBERS_ARRAY(new_first_name, new_memb_counter) & " " & NEW_MEMBERS_ARRAY(new_last_name, new_memb_counter)
			  Text 25, y_pos_1 + 20, 155, 10, "Relationship:" & right(NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter), len(NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_memb_counter)) - 5)
			  ' Text 25, 205, 155, 10, "Date of Change:"
			  ' Text 25, 215, 155, 10, "other"
			  pers_list_count = pers_list_count + 1
			  y_pos_1 = y_pos_1 + 40
		  End If
	  Next


	  GroupBox 190, 5, 185, 340, "Page 2"
	  Text 195, 20, 135, 20, "4. Do you want to apply for someone who is not getting coverage now?"
	  ' Text 205, 30, 125, 10, "is not getting coverage now?"
	  Text 340, 20, 35, 10, replace(apply_for_ma, "Did not answer", "BLANK")
	  y_pos_2 = 40
	  If q_4_details_blank_checkbox = checked then
		  Text 200, y_pos_2, 150, 10, "Q4 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
		  If NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb) <> "" Then
			  Text 200, y_pos_2, 150, 10, NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10

	  Text 195, y_pos_2, 135, 20, "5. Is anyone self-employed or does anyone expect to be self-employed?"
	  ' Text 205, 60, 125, 10, "anyone expect to be self-employed?"
	  Text 340, y_pos_2, 35, 10, replace(ma_self_employed, "Did not answer", "BLANK")
	  y_pos_2 = y_pos_2 + 20
	  If q_5_details_blank_checkbox = checked Then
		  Text 200, y_pos_2, 150, 10, "Q5 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
		  If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
			  Text 200, y_pos_2, 150, 10, NEW_EARNED_ARRAY(earned_client, each_busi) & " from " & NEW_EARNED_ARRAY(earned_source, each_busi) & " - $" & NEW_EARNED_ARRAY(earned_amount, each_busi) & " on " & NEW_EARNED_ARRAY(earned_start_date, each_busi)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10

	  Text 195, y_pos_2, 135, 20, "6. Does anyone work or does anyone expect to start working?"
	  ' Text 205, 90, 125, 10, "expect to start working?"
	  Text 340, y_pos_2, 35, 10, replace(ma_start_working, "Did not answer", "BLANK")
	  y_pos_2 = y_pos_2 + 20
	  If q_6_details_blank_checkbox = checked Then
		  Text 200, y_pos_2, 150, 10, "Q6 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
		  If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
			  Text 200, y_pos_2, 150, 10, NEW_EARNED_ARRAY(earned_client, each_job) & " from " & NEW_EARNED_ARRAY(earned_source, each_job) & " - $" & NEW_EARNED_ARRAY(earned_amount, each_job) & " on " & NEW_EARNED_ARRAY(earned_start_date, each_job)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10

	  Text 195, y_pos_2, 140, 25, "7. Does anyone get money or does anyone expect to get money from sources other than work?"
	  ' Text 205, 120, 130, 10, "anyone expect to get money from "
	  ' Text 205, 130, 115, 10, "sources other than work?"
	  Text 340, y_pos_2, 35, 10, replace(ma_other_income, "Did not answer", "BLANK")
	  y_pos_2 = y_pos_2 + 30
	  If q_7_details_blank_checkbox = checked Then
		  Text 200, y_pos_2, 150, 10, "Q7 details left BLANK"
		  y_pos_2 = y_pos_2 + 10
	  End If
	  For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
		  If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
			  Text 200, y_pos_2, 150, 10, NEW_UNEARNED_ARRAY(unearned_client, each_unea) & " from " & NEW_UNEARNED_ARRAY(unearned_source, each_unea) & " - $" & NEW_UNEARNED_ARRAY(unearned_amount, each_unea) & " on " & NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)
			  y_pos_2 = y_pos_2 + 10
		  End If
	  Next
	  y_pos_2 = y_pos_2 + 10


	  GroupBox 375, 5, 185, 340, "Page 3"
	  Text 380, 20, 145, 10, "9. Does anyone have cash or account?"
	  Text 525, 20, 35, 10, replace(ma_liquid_assets, "Did not answer", "BLANK")
	  y_pos_3 = 30
	  If q_9_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q9 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
		  Text 385, y_pos_3, 150, 10,  NEW_ASSET_ARRAY(asset_client, each_asset) & " Type -  " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
		  y_pos_3 = y_pos_3 + 10
		End If
	  Next
	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 135, 20, "10. Does anyone own securities or other assets?"
	  ' Text 390, 50, 125, 10, "other assets?"
	  Text 525, y_pos_3, 35, 10, replace(ma_security_assets, "Did not answer", "BLANK")
	  y_pos_3 = y_pos_3 + 20
	  If q_10_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q10 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			Text 385, y_pos_3, 150, 10,  NEW_ASSET_ARRAY(asset_client, each_asset) & " Type -  " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
			y_pos_3 = y_pos_3 + 10
		End If
	  Next
	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 135, 10, "11. Does anyone own a vehicle?"
	  Text 525, y_pos_3, 35, 10, replace(ma_vehicle, "Did not answer", "BLANK")
	  y_pos_3 = y_pos_3 + 10
	  If q_11_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q11 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
		  If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
			  Text 385, y_pos_3, 150, 10,  NEW_ASSET_ARRAY(asset_client, each_asset) & " Type -  " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
			  y_pos_3 = y_pos_3 + 10
		  End If
	  Next
	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 140, 20, "12. Does anyone own or co-own a house or any real estate?"
	  ' Text 390, 100, 130, 10, "or any real estate?"
	  Text 525, y_pos_3, 35, 10, replace(ma_real_assets, "Did not answer", "BLANK")
	  y_pos_3 = y_pos_3 + 20
	  If q_12_details_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q12 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If

	  y_pos_3 = y_pos_3 + 10

	  Text 380, y_pos_3, 140, 10, "13. Do you have any change to report?"
	  Text 525, y_pos_3, 35, 10, replace(ma_other_changes, "Did not answer", "BLANK")
	  y_pos_3 = y_pos_3 + 10
	  If changes_reported_blank_checkbox = checked Then
		  Text 385, y_pos_3, 150, 10, "Q13 details left BLANK"
		  y_pos_3 = y_pos_3 + 10
	  End If
	  If trim(other_changes_reported) <> "" Then
		  Text 385, y_pos_3, 150, 10, "Other changes: " & other_changes_reported
		  y_pos_3 = y_pos_3 + 10
	  ENd If

	  y_pos_3 = y_pos_3 + 10


	  GroupBox 560, 5, 185, 340, "Page 4"
	  Text 585, 20, 135, 20, "Since your last application or in the past six months..."
	  Text 565, 45, 125, 10, "15. Has your household moved?"
	  Text 710, 45, 35, 10, replace(quest_fifteen_form_answer, "Did not answer", "BLANK")
	  y_pos_4 = 55
	  If trim(new_rent_or_mortgage_amount) = "" AND heat_ac_checkbox = unchecked AND electricity_checkbox = unchecked AND telephone_checkbox = unchecked Then
		  Text 570, y_pos_4, 150, 10, "Q15 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  Else
		  If trim(new_rent_or_mortgage_amount) = "" Then Text 570, y_pos_4, 150, 10, "NO new shelter Cost"
		  If trim(new_rent_or_mortgage_amount) <> "" THen Text 570, y_pos_4, 150, 10, "New Shelter Cost: $" & new_rent_or_mortgage_amount
		  y_pos_4 = y_pos_4 + 10

		  If heat_ac_checkbox = checked OR electricity_checkbox = checked OR telephone_checkbox = checked Then
			  Text 570, y_pos_4, 50, 10, "Utilities Paid"
			  y_pos_4 = y_pos_4 + 10
			  If heat_ac_checkbox = checked Then Text 575, y_pos_4, 50, 10, "HEAT/AC"
			  If electricity_checkbox = checked Then Text 625, y_pos_4, 50, 10, "ELECTRIC"
			  If telephone_checkbox = checked Then Text 675, y_pos_4, 50, 10, "PHONE"
			  y_pos_4 = y_pos_4 + 10
		  End If

		  ' Text 570, y_pos_4, 150, 10, dlg_text

	  End If

	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 20, "16. Has anyone had a change in their income from work?"
	  ' Text 575, 75, 125, 10, ""
	  Text 710, y_pos_4, 35, 10, replace(quest_sixteen_form_answer, "Did not answer", "BLANK")
	  y_pos_4 = y_pos_4 + 20
	  If q_16_details_blank_checkbox = checked Then
		  Text 570, y_pos_4, 150, 10, "Q16 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  End If
	  For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
		  If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
			  Text 570, y_pos_4, 150, 10, NEW_EARNED_ARRAY(earned_client, the_earned) & " from " & NEW_EARNED_ARRAY(earned_source, the_earned) & " - $" & NEW_EARNED_ARRAY(earned_amount, the_earned) & " on " & NEW_EARNED_ARRAY(earned_change_date, the_earned)

			  y_pos_4 = y_pos_4 + 10
		  End If
	  Next
	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 140, 25, "17. Has anyone had a change of more than $50 per month from income sources other than work or a change in unearned income?"
	  ' Text 575, 105, 140, 10, "than $50 per month from income sources"
	  ' Text 575, 115, 160, 10, "other than work or a change in unearned income?"
	  Text 710, y_pos_4, 35, 10, replace(quest_seventeen_form_answer, "Did not answer", "BLANK")
	  y_pos_4 = y_pos_4 + 30
	  If q_17_details_blank_checkbox = checked Then
		  Text 570, y_pos_4, 150, 10, "Q17 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  End If
	  For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
		  If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
			  Text 570, y_pos_4, 150, 10, NEW_UNEARNED_ARRAY(unearned_client, the_unearned) & " from " & NEW_UNEARNED_ARRAY(unearned_source, the_unearned) & " - $" & NEW_UNEARNED_ARRAY(unearned_amount, the_unearned) & " on " & NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned)

			  y_pos_4 = y_pos_4 + 10
		  End If
	  Next
	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 25, "18. Has anyone had a change in court-ordered child or medical support payments?"
	  ' Text 575, 145, 140, 10, "court-ordered child or medical "
	  ' Text 575, 155, 160, 10, "support payments?"
	  Text 710, y_pos_4, 35, 10, replace(quest_eighteen_form_answer, "Did not answer", "BLANK")
	  y_pos_4 = y_pos_4 + 30
	  If q_18_details_blank_checkbox = checked Then
		  Text 570, y_pos_4, 150, 10, "Q18 details left BLANK"
		  y_pos_4 = y_pos_4 + 10
	  End If
	  For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
		  If NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "" THen
			  Text 570, y_pos_4, 150, 10, "Child support - paid by: " & NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) & " - $" & NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs)

			  y_pos_4 = y_pos_4 + 10
		  End If
	  Next
	  y_pos_4 = y_pos_4 + 10

	  Text 565, y_pos_4, 135, 25, "19. Did you work 20 hours each week, for an average of 80 hours each month during the past six months?"
	  ' Text 575, 185, 140, 10, "for an average of 80 hours each month"
	  ' Text 575, 195, 160, 10, "during the past six months?"
	  Text 710, y_pos_4, 35, 10, replace(quest_nineteen_form_answer, "Did not answer", "BLANK")
	  GroupBox 560, 285, 185, 60, "Page 5"
	  Text 570, 300, 165, 10, "Signature:" & client_signed_yn
	  Text 570, 315, 165, 10, "Date:" & client_dated_yn

	  Text 10, 355, 240, 10, "Review the information here, does it match the form the client submited?"

	EndDialog

	dialog Dialog1
	cancel_confirmation

	err_msg = "LOOP"

	If ButtonPressed = back_to_dlg_addr Then
		show_csr_dlg_q_1 = TRUE
		show_csr_dlg_q_2 = TRUE
	End If
	If ButtonPressed = back_to_dlg_ma_income Then show_csr_dlg_q_4_7 = TRUE
	If ButtonPressed = back_to_dlg_ma_asset Then
		show_csr_dlg_q_9_12 = TRUE
		show_csr_dlg_q_13 = TRUE
	End If
	If ButtonPressed = back_to_dlg_snap Then show_csr_dlg_q_15_19 = TRUE
	If ButtonPressed = back_to_dlg_sig Then show_csr_dlg_sig = TRUE
	' MsgBox show_csr_dlg_q_15_19
	If ButtonPressed = -1 Then
		err_msg = ""
		If confirm_csr_form_information = "Indicate the form information" THen err_msg = err_msg & vbNewLine & "* Indicate if this information is correct and matches the form received. If something is not correct, use the buttons on this dialog to go back to the correct area and update the information on the specific dialog."
		If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please Resolve to Continue:" & vbNewLine & err_msg
		If confirm_csr_form_information = "NO - the information here is different" Then
			show_csr_dlg_q_1 = TRUE
			show_csr_dlg_q_2 = TRUE
			show_csr_dlg_q_4_7 = TRUE
			show_csr_dlg_q_9_12 = TRUE
			show_csr_dlg_q_13 = TRUE
			show_csr_dlg_q_15_19 = TRUE
			show_csr_dlg_sig = TRUE
		End If
	Else
		confirm_csr_form_information = "Indicate the form information"
	End If
end function

function gather_pers_detail()
	If grh_sr_yn = "Yes" Then grh_sr = TRUE
	If hc_sr_yn = "Yes" Then hc_sr = TRUE
	If snap_sr_yn = "Yes" Then snap_sr = TRUE

	If grh_sr = TRUE Then
	    MAXIS_footer_month = grh_sr_mo
	    MAXIS_footer_year = grh_sr_yr
	End If
	If hc_sr = TRUE Then
	    MAXIS_footer_month = hc_sr_mo
	    MAXIS_footer_year = hc_sr_yr
	End If
	If snap_sr = TRUE Then
	    MAXIS_footer_month = snap_sr_mo
	    MAXIS_footer_year = snap_sr_yr
	End If

	Call navigate_to_MAXIS_screen("CASE", "PERS")

	pers_row = 10                                               'This is where client information starts on CASE PERS
	person_counter = 0
	Do
	    EMReadScreen the_snap_status, 1, pers_row, 54
	    EMReadScreen the_grh_status, 1, pers_row, 66
	    EMReadScreen the_hc_status, 1, pers_row, 61             'reading the HC status of each client
	    ' MsgBox the_snap_status & vbNewLine & person_counter
	    If the_snap_status = "A" Then
	        ALL_CLIENTS_ARRAY(clt_snap_status, person_counter) = "Active"
	    ElseIf the_snap_status = "P" Then
	        ALL_CLIENTS_ARRAY(clt_snap_status, person_counter) = "Pending"
	    Else
	        ALL_CLIENTS_ARRAY(clt_snap_status, person_counter) = "Inactive"
	    End If
	    If the_grh_status = "A" Then
	        ALL_CLIENTS_ARRAY(clt_grh_status, person_counter) = "Active"
	    ElseIf the_grh_status = "P" Then
	        ALL_CLIENTS_ARRAY(clt_grh_status, person_counter) = "Pending"
	    Else
	        ALL_CLIENTS_ARRAY(clt_grh_status, person_counter) = "Inactive"
	    End If
	    If the_hc_status = "A" Then
	        ALL_CLIENTS_ARRAY(clt_hc_status, person_counter) = "Active"
	    ElseIf the_hc_status = "P" Then
	        ALL_CLIENTS_ARRAY(clt_hc_status, person_counter) = "Pending"
	    Else
	        ALL_CLIENTS_ARRAY(clt_hc_status, person_counter) = "Inactive"
	    End If

	    person_counter = person_counter + 1
	    pers_row = pers_row + 3         'next client information is 3 rows down
	    If pers_row = 19 Then           'this is the end of the list of client on each list
	        PF8                         'going to the next page of client information
	        pers_row = 10
	        EmReadscreen end_of_list, 9, 24, 14
	        If end_of_list = "LAST PAGE" Then Exit Do
	    End If
	    EmReadscreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list

	Loop until next_pers_ref_numb = "  "
	Call back_to_SELF

	Call navigate_to_MAXIS_screen("STAT", "WREG")

	For all_the_membs = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
		If ALL_CLIENTS_ARRAY(clt_snap_status, all_the_membs) = "Active" OR ALL_CLIENTS_ARRAY(clt_snap_status, all_the_membs) = "Pending" Then
			EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, all_the_membs), 20, 76
			transmit

			EMReadScreen wreg_abawd_code, 2, 13, 50
			If wreg_abawd_code = "09" OR wreg_abawd_code = "10" OR wreg_abawd_code = "11" OR wreg_abawd_code = "13" Then abawd_on_case = TRUE
		End If
	Next
end function

function count_actives()
	snap_active_count = 0
	hc_active_count = 0
	grh_active_count = 0
	For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
	    If ALL_CLIENTS_ARRAY(clt_grh_status, known_memb) = "Active" Then grh_active_count = grh_active_count + 1
	    If ALL_CLIENTS_ARRAY(clt_hc_status, known_memb) = "Active" Then hc_active_count = hc_active_count + 1
	    If ALL_CLIENTS_ARRAY(clt_snap_status, known_memb) = "Active" Then snap_active_count = snap_active_count + 1
	Next
end function

function enter_new_residence_address()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 356, 135, "New Residence Address Information"
		  Text 240, 25, 50, 10, "Effective Date"
		  EditBox 300, 20, 50, 15, new_addr_effective_date
		  Text 10, 25, 145, 10, "New Residence Address Reported on CSR:"
		  Text 20, 45, 45, 10, "House/Street:"
		  EditBox 70, 40, 280, 15, new_resi_one
		  Text 50, 65, 15, 10, "City:"
		  EditBox 70, 60, 80, 15, new_resi_city
		  Text 160, 65, 20, 10, "State:"
		  DropListBox 185, 60, 75, 45, state_list, new_resi_state
		  Text 275, 65, 20, 10, "Zip:"
		  EditBox 300, 60, 50, 15, new_resi_zip
		  Text 40, 85, 30, 10, "County:"
		  DropListBox 70, 80, 190, 45, "Select One..."+chr(9)+county_list, new_resi_county
		  Text 95, 100, 90, 10, "Address/Home Verification:"
		  DropListBox 190, 95, 125, 45, "Select One..."+chr(9)+"SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Ver Prvd", new_shel_verif
		  Text 10, 5, 300, 10, "ENTER THE RESIDENCE ADDRESS INFORMATION FROM THE CSR FORM"
		  ButtonGroup ButtonPressed
		    OkButton 300, 115, 50, 15
		EndDialog

		dialog Dialog1

		If trim(new_addr_effective_date) <> "" AND IsDate(new_addr_effective_date) = FALSE THen err_msg = err_msg & vbNewLine & "* Enter the effective date as a valid date or leave blank."
		new_resi_one = trim(new_resi_one)
		new_resi_city = trim(new_resi_city)
		new_resi_zip = trim(new_resi_zip)
		If new_resi_one = "" AND new_resi_city = "" AND new_resi_state = "Select One..." AND new_resi_zip = "" Then err_msg = err_msg & vbNewLine & "* Enter the details from the form."
		If new_resi_county = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the county of residence."
		If new_shel_verif = "Select One..." Then err_msg = err_msg & vbNewLine & "* Enter the verification (NO or OT are acceptable)."

		If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbNewLine & "Please Resolve to Continue:" & vbNewLine & err_msg

	Loop until err_msg = ""
	residence_address_match_yn = "No - New Address Entered"
	new_resi_addr_entered = TRUE
end function

function enter_new_mailing_address()
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 356, 95, "New Mailing Address Information"
		  Text 10, 20, 145, 10, "New Mailing Address Reported on CSR:"
		  Text 20, 40, 45, 10, "House/Street:"
		  EditBox 70, 35, 280, 15, new_mail_one
		  Text 50, 60, 15, 10, "City:"
		  EditBox 70, 55, 80, 15, new_mail_city
		  Text 160, 60, 20, 10, "State:"
		  DropListBox 185, 55, 75, 45, state_list, new_mail_state
		  Text 275, 60, 20, 10, "Zip:"
		  EditBox 300, 55, 50, 15, new_mail_zip
		  Text 10, 5, 300, 10, "ENTER THE MAILING ADDRESS INFORMATION FROM THE CSR FORM"
		  ButtonGroup ButtonPressed
		    OkButton 300, 75, 50, 15
		EndDialog

		dialog Dialog1

		new_mail_one = trim(new_mail_one)
		new_mail_city = trim(new_mail_city)
		new_mail_zip = trim(new_mail_zip)

		If new_mail_one = "" AND new_mail_city = "" AND new_mail_state = "Select One..." AND new_mail_zip = "" Then err_msg = err_msg & vbNewLine & "* Enter the details from the form."
		If err_msg <> "" Then MsgBox "****** NOTICE ******" & vbNewLine & "Please Resolve to Continue:" & vbNewLine & err_msg
	Loop until err_msg = ""
	mailing_address_match_yn = "No - New Address Entered"
	new_mail_addr_entered = TRUE
end function

function get_county_name_from_county_code(county_code, county_name, include_county_code)
    If county_code = "01" Then county_name = "Aitkin"
    If county_code = "02" Then county_name = "Anoka"
    If county_code = "03" Then county_name = "Becker"
    If county_code = "04" Then county_name = "Beltrami"
    If county_code = "05" Then county_name = "Benton"
    If county_code = "06" Then county_name = "Big Stone"
    If county_code = "07" Then county_name = "Blue Earth"
    If county_code = "08" Then county_name = "Brown"
    If county_code = "09" Then county_name = "Carlton"
    If county_code = "10" Then county_name = "Carver"
    If county_code = "11" Then county_name = "Cass"
    If county_code = "12" Then county_name = "Chippewa"
    If county_code = "13" Then county_name = "Chisago"
    If county_code = "14" Then county_name = "Clay"
    If county_code = "15" Then county_name = "Clearwater"
    If county_code = "16" Then county_name = "Cook"
    If county_code = "17" Then county_name = "Cottonwood"
    If county_code = "18" Then county_name = "Crow Wing"
    If county_code = "19" Then county_name = "Dakota"
    If county_code = "20" Then county_name = "Dodge"
    If county_code = "21" Then county_name = "Douglas"
    If county_code = "22" Then county_name = "Faribault"
    If county_code = "23" Then county_name = "Fillmore"
    If county_code = "24" Then county_name = "Freeborn"
    If county_code = "25" Then county_name = "Goodhue"
    If county_code = "26" Then county_name = "Grant"
    If county_code = "27" Then county_name = "Hennepin"
    If county_code = "28" Then county_name = "Houston"
    If county_code = "29" Then county_name = "Hubbard"
    If county_code = "30" Then county_name = "Isanti"
    If county_code = "31" Then county_name = "Itasca"
    If county_code = "32" Then county_name = "Jackson"
    If county_code = "33" Then county_name = "Kanabec"
    If county_code = "34" Then county_name = "Kandiyohi"
    If county_code = "35" Then county_name = "Kittson"
    If county_code = "36" Then county_name = "Koochiching"
    If county_code = "37" Then county_name = "Lac Qui Parle"
    If county_code = "38" Then county_name = "Lake"
    If county_code = "39" Then county_name = "Lake Of Woods"
    If county_code = "40" Then county_name = "Le Sueur"
    If county_code = "41" Then county_name = "Lincoln"
    If county_code = "42" Then county_name = "Lyon"
    If county_code = "43" Then county_name = "Mcleod"
    If county_code = "44" Then county_name = "Mahnomen"
    If county_code = "45" Then county_name = "Marshall"
    If county_code = "46" Then county_name = "Martin"
    If county_code = "47" Then county_name = "Meeker"
    If county_code = "48" Then county_name = "Mille Lacs"
    If county_code = "49" Then county_name = "Morrison"
    If county_code = "50" Then county_name = "Mower"
    If county_code = "51" Then county_name = "Murray"
    If county_code = "52" Then county_name = "Nicollet"
    If county_code = "53" Then county_name = "Nobles"
    If county_code = "54" Then county_name = "Norman"
    If county_code = "55" Then county_name = "Olmsted"
    If county_code = "56" Then county_name = "Otter Tail"
    If county_code = "57" Then county_name = "Pennington"
    If county_code = "58" Then county_name = "Pine"
    If county_code = "59" Then county_name = "Pipestone"
    If county_code = "60" Then county_name = "Polk"
    If county_code = "61" Then county_name = "Pope"
    If county_code = "62" Then county_name = "Ramsey"
    If county_code = "63" Then county_name = "Red Lake"
    If county_code = "64" Then county_name = "Redwood"
    If county_code = "65" Then county_name = "Renville"
    If county_code = "66" Then county_name = "Rice"
    If county_code = "67" Then county_name = "Rock"
    If county_code = "68" Then county_name = "Roseau"
    If county_code = "69" Then county_name = "St. Louis"
    If county_code = "70" Then county_name = "Scott"
    If county_code = "71" Then county_name = "Sherburne"
    If county_code = "72" Then county_name = "Sibley"
    If county_code = "73" Then county_name = "Stearns"
    If county_code = "74" Then county_name = "Steele"
    If county_code = "75" Then county_name = "Stevens"
    If county_code = "76" Then county_name = "Swift"
    If county_code = "77" Then county_name = "Todd"
    If county_code = "78" Then county_name = "Traverse"
    If county_code = "79" Then county_name = "Wabasha"
    If county_code = "80" Then county_name = "Wadena"
    If county_code = "81" Then county_name = "Waseca"
    If county_code = "82" Then county_name = "Washington"
    If county_code = "83" Then county_name = "Watonwan"
    If county_code = "84" Then county_name = "Wilkin"
    If county_code = "85" Then county_name = "Winona"
    If county_code = "86" Then county_name = "Wright"
    If county_code = "87" Then county_name = "Yellow Medicine"
    If county_code = "89" Then county_name = "Out-of-State"

    If include_county_code = TRUE Then county_name = county_code & " " & county_name
end function

function get_state_name_from_state_code(state_code, state_name, include_state_code)
    If state_code = "NB" Then state_name = "MN Newborn"
    If state_code = "FC" Then state_name = "Foreign Country"
    If state_code = "UN" Then state_name = "Unknown"
    If state_code = "AL" Then state_name = "Alabama"
    If state_code = "AK" Then state_name = "Alaska"
    If state_code = "AZ" Then state_name = "Arizona"
    If state_code = "AR" Then state_name = "Arkansas"
    If state_code = "CA" Then state_name = "California"
    If state_code = "CO" Then state_name = "Colorado"
    If state_code = "CT" Then state_name = "Connecticut"
    If state_code = "DE" Then state_name = "Delaware"
    If state_code = "DC" Then state_name = "District Of Columbia"
    If state_code = "FL" Then state_name = "Florida"
    If state_code = "GA" Then state_name = "Georgia"
    If state_code = "HI" Then state_name = "Hawaii"
    If state_code = "ID" Then state_name = "Idaho"
    If state_code = "IL" Then state_name = "Illnois"
    If state_code = "IN" Then state_name = "Indiana"
    If state_code = "IA" Then state_name = "Iowa"
    If state_code = "KS" Then state_name = "Kansas"
    If state_code = "KY" Then state_name = "Kentucky"
    If state_code = "LA" Then state_name = "Louisiana"
    If state_code = "ME" Then state_name = "Maine"
    If state_code = "MD" Then state_name = "Maryland"
    If state_code = "MA" Then state_name = "Massachusetts"
    If state_code = "MI" Then state_name = "Michigan"
    If state_code = "MS" Then state_name = "Mississippi"
    If state_code = "MO" Then state_name = "Missouri"
    If state_code = "MT" Then state_name = "Montana"
    If state_code = "NE" Then state_name = "Nebraska"
    If state_code = "NV" Then state_name = "Nevada"
    If state_code = "NH" Then state_name = "New Hampshire"
    If state_code = "NJ" Then state_name = "New Jersey"
    If state_code = "NM" Then state_name = "New Mexico"
    If state_code = "NY" Then state_name = "New York"
    If state_code = "NC" Then state_name = "North Carolina"
    If state_code = "ND" Then state_name = "North Dakota"
    If state_code = "OH" Then state_name = "Ohio"
    If state_code = "OK" Then state_name = "Oklahoma"
    If state_code = "OR" Then state_name = "Oregon"
    If state_code = "PA" Then state_name = "Pennsylvania"
    If state_code = "RI" Then state_name = "Rhode Island"
    If state_code = "SC" Then state_name = "South Carolina"
    If state_code = "SD" Then state_name = "South Dakota"
    If state_code = "TN" Then state_name = "Tennessee"
    If state_code = "TX" Then state_name = "Texas"
    If state_code = "UT" Then state_name = "Utah"
    If state_code = "VT" Then state_name = "Vermont"
    If state_code = "VA" Then state_name = "Virginia"
    If state_code = "WA" Then state_name = "Washington"
    If state_code = "WV" Then state_name = "West Virginia"
    If state_code = "WI" Then state_name = "Wisconsin"
    If state_code = "WY" Then state_name = "Wyoming"
    If state_code = "PR" Then state_name = "Puerto Rico"
    If state_code = "VI" Then state_name = "Virgin Islands"

    If include_state_code = TRUE Then state_name = state_code & " " & state_name
end function

function get_state_code_from_state_name(state_name, state_code)
    If state_name = "Alabama"           Then state_code = "AL"
    If state_name = "Alaska"            Then state_code = "AK"
    If state_name = "Arizona"           Then state_code = "AZ"
    If state_name = "Arkansas"          Then state_code = "AR"
    If state_name = "California"        Then state_code = "CA"
    If state_name = "Colorado"          Then state_code = "CO"
    If state_name = "Connecticut"       Then state_code = "CT"
    If state_name = "Delaware"          Then state_code = "DE"
    If state_name = "Florida"           Then state_code = "FL"
    If state_name = "Georgia"           Then state_code = "GA"
    If state_name = "Hawaii"            Then state_code = "HI"
    If state_name = "Idaho"             Then state_code = "ID"
    If state_name = "Illinois"          Then state_code = "IL"
    If state_name = "Indiana"           Then state_code = "IN"
    If state_name = "Iowa"              Then state_code = "IA"
    If state_name = "Kansas"            Then state_code = "KS"
    If state_name = "Kentucky"          Then state_code = "KY"
    If state_name = "Louisiana"         Then state_code = "LA"
    If state_name = "Maine"             Then state_code = "ME"
    If state_name = "Maryland"          Then state_code = "MD"
    If state_name = "Massachusetts"     Then state_code = "MA"
    If state_name = "Michigan"          Then state_code = "MI"
    If state_name = "Mississippi"       Then state_code = "MS"
    If state_name = "Missouri"          Then state_code = "MO"
    If state_name = "Montana"           Then state_code = "MT"
    If state_name = "Nebraska"          Then state_code = "NE"
    If state_name = "Nevada"            Then state_code = "NV"
    If state_name = "New Hampshire"     Then state_code = "NH"
    If state_name = "New Jersey"        Then state_code = "NJ"
    If state_name = "New Mexico"        Then state_code = "NM"
    If state_name = "New York"          Then state_code = "NY"
    If state_name = "North Carolina"    Then state_code = "NC"
    If state_name = "North Dakota"      Then state_code = "ND"
    If state_name = "Ohio"              Then state_code = "OH"
    If state_name = "Oklahoma"          Then state_code = "OK"
    If state_name = "Oregon"            Then state_code = "OR"
    If state_name = "Pennsylvania"      Then state_code = "PA"
    If state_name = "Rhode Island"      Then state_code = "RI"
    If state_name = "South Carolina"    Then state_code = "SC"
    If state_name = "South Dakota"      Then state_code = "SD"
    If state_name = "Tennessee"         Then state_code = "TN"
    If state_name = "Texas"             Then state_code = "TX"
    If state_name = "Utah"              Then state_code = "UT"
    If state_name = "Vermont"           Then state_code = "VT"
    If state_name = "Virginia"          Then state_code = "VA"
    If state_name = "Washington"        Then state_code = "WA"
    If state_name = "West Virginia"     Then state_code = "WV"
    If state_name = "Wisconsin"         Then state_code = "WI"
    If state_name = "Wyoming"           Then state_code = "WY"
end function


function validate_footer_month_entry(footer_month, footer_year, err_msg_var, bullet_char)
    If IsNumeric(footer_month) = FALSE Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be a number, review and reenter the footer month information."
    Else
        footer_month = footer_month * 1
        If footer_month > 12 OR footer_month < 1 Then err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be between 1 and 12. Review and reenter the footer month information."
        footer_month = right("00" & footer_month, 2)
    End If

    If len(footer_year) < 2 Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be at least 2 characters long, review and reenter the footer year information."
    Else
        If IsNumeric(footer_year) = FALSE Then
            err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be a number, review and reenter the footer year information."
        Else
            footer_year = right("00" & footer_year, 2)
        End If
    End If
end function
'===========================================================================================================================
'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim MAXIS_footer_month, MAXIS_footer_year, snap_active_count, hc_active_count, grh_active_count, snap_sr_yn, snap_sr_mo, snap_sr_yr, hc_sr_yn, hc_sr_mo, hc_sr_yr, grh_sr_yn, grh_sr_mo, grh_sr_yr, client_on_csr_form
Dim residence_address_match_yn, mailing_address_match_yn, homeless_status, grh_sr, hc_sr, snap_sr, notes_on_address, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, new_mail_zip
Dim quest_two_move_in_out, new_hh_memb_not_in_mx_yn, apply_for_ma, q_4_details_blank_checkbox, ma_self_employed, q_5_details_blank_checkbox, ma_start_working, q_6_details_blank_checkbox, ma_other_income
Dim q_7_details_blank_checkbox, ma_liquid_assets, q_9_details_blank_checkbox, ma_security_assets, q_10_details_blank_checkbox, ma_vehicle, q_11_details_blank_checkbox, ma_real_assets, q_12_details_blank_checkbox
Dim ma_other_changes, other_changes_reported, changes_reported_blank_checkbox, quest_fifteen_form_answer, new_rent_or_mortgage_amount, heat_ac_checkbox, electricity_checkbox, telephone_checkbox, shel_proof_provided
Dim quest_sixteen_form_answer, q_16_details_blank_checkbox, quest_seventeen_form_answer, q_17_details_blank_checkbox, quest_eighteen_form_answer, q_18_details_blank_checkbox, quest_nineteen_form_answer, csr_form_date
Dim addr_verif, addr_homeless, addr_reservation, living_situation_status, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, new_mail_one, new_mail_city, new_mail_state
Dim client_signed_yn, client_dated_yn, confirm_csr_form_information, notes_on_faci, notes_on_wreg, new_addr_effective_date, new_resi_one, new_resi_city, new_resi_state, new_resi_zip, new_resi_county, new_shel_verif
Dim new_resi_addr_entered, new_mail_addr_entered

HH_memb_row = 5
Dim row
Dim col

Const owner_name                = 00
Const category_const            = 01
Const type_const                = 02
Const name_const                = 03
Const amount_const              = 04
Const start_date_const          = 05
Const end_date_const            = 06
Const verif_const               = 07
Const pay_amt_const             = 08
Const hours_const               = 09
Const update_date_const         = 10
Const seasonal_yn               = 11
Const frequency_const           = 12
Const make_const                = 13
Const model_const               = 14
Const year_const                = 15
Const make_model_yr             = 16
Const address_const             = 17
Const cash_amt_const            = 18        'cash panel
Const cash_verif_const          = 19
Const snap_amt_const            = 20
Const snap_verif_const          = 21
Const hc_amt_const              = 22
Const hc_verif_const            = 23
Const busi_cash_net_prosp       = 24        'busi panel
Const busi_cash_net_retro       = 25
Const busi_cash_gross_retro     = 26
Const busi_cash_expense_retro   = 27
Const busi_cash_gross_prosp     = 28
Const busi_cash_expense_prosp   = 29
Const busi_cash_income_verif    = 30
Const busi_cash_expense_verif   = 31
Const busi_snap_net_prosp       = 32
Const busi_snap_net_retro       = 33
Const busi_snap_gross_retro     = 34
Const busi_snap_expense_retro   = 35
Const busi_snap_gross_prosp     = 36
Const busi_snap_expense_prosp   = 37
Const busi_snap_income_verif    = 38
Const busi_snap_expense_verif   = 39
Const busi_hc_a_net_prosp       = 40
Const busi_hc_a_gross_prosp     = 41
Const busi_hc_a_expense_prosp   = 42
Const busi_hc_a_income_verif    = 43
Const busi_hc_a_expense_verif   = 44
Const busi_hc_b_net_prosp       = 45
Const busi_hc_b_gross_prosp     = 46
Const busi_hc_b_expense_prosp   = 47
Const busi_hc_b_income_verif    = 48
Const busi_hc_b_expense_verif   = 49
Const busi_se_method            = 50
Const busi_se_method_date       = 51
Const rptd_hours_const          = 52
Const min_wg_hours_const        = 53
Const claim_nbr_const           = 54        'unea panel
Const cola_disregard_amt        = 55
Const id_number_const           = 56
Const panel_instance            = 57
Const owner_ref_const           = 58
Const verif_checkbox_const      = 59
Const verif_time_const          = 60
Const verif_added_const         = 61
Const item_notes_const          = 62
Const balance_date_const        = 63
Const withdraw_penalty_const    = 64
Const withdraw_yn_const         = 65
Const withdraw_verif_const      = 66
Const count_cash_const          = 67
Const count_snap_const          = 68
Const count_hc_const            = 69
Const count_grh_const           = 70
Const count_ive_const           = 71
Const joint_own_const           = 72
Const share_ratio_const         = 73
Const next_interst_const        = 74
Const face_value_const          = 75
Const trade_in_const            = 76
Const loan_const                = 77
Const source_const              = 78
Const owed_amt_const            = 79
Const owed_verif_const          = 80
Const owed_date_const           = 81
Const cars_use_const            = 82
Const hc_benefit_const          = 83
Const market_value_const        = 84
Const value_verif_const         = 85
Const rest_prop_status_const    = 86
Const rest_repymt_date_const    = 87

Const jobs_hrly_wage            = 88
Const retro_income_amount       = 89
Const retro_income_hours        = 90
Const snap_pic_frequency        = 91
Const snap_pic_hours_per_pay    = 92
Const snap_pic_income_per_pay   = 93
Const snap_pic_monthly_income   = 94
Const grh_pic_frequency         = 95
Const grh_pic_income_per_pay    = 96
Const grh_pic_monthly_income    = 97
Const jobs_subsidy              = 98

Const new_checkbox              = 99
Const update_checkbox           = 100

Const faci_ref_numb                 = 00
Const faci_instance                 = 01
Const faci_member                   = 02
Const faci_name                     = 03
Const faci_vendor_number            = 04
Const faci_type                     = 05
Const faci_FS_elig                  = 06
Const faci_FS_type                  = 07
Const faci_waiver_type              = 08
Const faci_ltc_inelig_reason        = 09
Const faci_inelig_begin_date        = 10
Const faci_inelig_end_date          = 11
Const faci_anticipated_out_date     = 12
Const faci_GRH_plan_required        = 13
Const faci_GRH_plan_verif           = 14
Const faci_cty_app_place            = 15
Const faci_approval_cty_name        = 16
Const faci_GRH_DOC_amount           = 17
Const faci_GRH_postpay              = 18
Const faci_stay_one_rate            = 19
Const faci_stay_one_date_in         = 20
Const faci_stay_one_date_out        = 21
Const faci_stay_two_rate            = 22
Const faci_stay_two_date_in         = 23
Const faci_stay_two_date_out        = 24
Const faci_stay_three_rate          = 25
Const faci_stay_three_date_in       = 26
Const faci_stay_three_date_out      = 27
Const faci_stay_four_rate           = 28
Const faci_stay_four_date_in        = 29
Const faci_stay_four_date_out       = 30
Const faci_stay_five_rate           = 31
Const faci_stay_five_date_in        = 32
Const faci_stay_five_date_out       = 33
Const faci_verif_checkbox           = 34
Const faci_verif_added              = 35
Const faci_notes                    = 36

Dim ALL_INCOME_ARRAY()
ReDim ALL_INCOME_ARRAY(update_checkbox, 0)

Dim ALL_ASSETS_ARRAY()
ReDim ALL_ASSETS_ARRAY(update_checkbox, 0)

Dim FACILITIES_ARRAY()
ReDim FACILITIES_ARRAY(faci_notes, 0)

const memb_ref_numb                 = 00
const memb_last_name                = 01
const memb_first_name               = 02
const memb_age                      = 03
const memb_remo_checkbox            = 04
const memb_new_checkbox             = 05
const clt_grh_status                = 06
const clt_hc_status                 = 07
const clt_snap_status               = 08
const memb_id_verif                 = 09
const memb_soc_sec_numb             = 10
const memb_ssn_verif                = 11
const memb_dob                      = 12
const memb_dob_verif                = 13
const memb_gender                   = 14
const memb_rel_to_applct            = 15
const memb_spoken_language          = 16
const memb_written_language         = 17
const memb_interpreter              = 18
const memb_alias                    = 19
const memb_ethnicity                = 20
const memb_race                     = 21
const memi_marriage_status          = 22
const memi_spouse_ref               = 23
const memi_spouse_name              = 24
const memi_designated_spouse        = 25
const memi_marriage_date            = 26
const memi_marriage_verif           = 27
const memi_citizen                  = 28
const memi_citizen_verif            = 29
const memi_last_grade               = 30
const memi_in_MN_less_12_mo         = 31
const memi_resi_verif               = 32
const memi_MN_entry_date            = 33
const memi_former_state             = 34
const memi_other_FS_end             = 35
const wreg_pwe                      = 36
const wreg_status                   = 37
const wreg_defer_fset               = 38
const wreg_fset_orient_date         = 39
const wreg_sanc_begin_date          = 40
const wreg_sanc_count               = 41
const wreg_sanc_reasons             = 42
const wreg_abawd_status             = 43
const wreg_banekd_months            = 44
const wreg_GA_basis                 = 45
const wreg_GA_coop                  = 46
const wreg_numb_ABAWD_months        = 47
const wreg_ABAWD_months_list        = 48
const wreg_numb_second_set_months   = 49
const wreg_second_set_months_list   = 50
Const wreg_notes                    = 51

const shel_hud_sub_yn               = 52
const shel_shared_yn                = 53
const shel_paid_to                  = 54
const shel_rent_retro_amt           = 55
const shel_rent_retro_verif         = 56
const shel_rent_prosp_amt           = 57
const shel_rent_prosp_verif         = 58
const shel_lot_rent_retro_amt       = 59
const shel_lot_rent_retro_verif     = 60
const shel_lot_rent_prosp_amt       = 61
const shel_lot_rent_prosp_verif     = 62
const shel_mortgage_retro_amt       = 63
const shel_mortgage_retro_verif     = 64
const shel_mortgage_prosp_amt       = 65
const shel_mortgage_prosp_verif     = 66
const shel_insurance_retro_amt      = 67
const shel_insurance_retro_verif    = 68
const shel_insurance_prosp_amt      = 69
const shel_insurance_prosp_verif    = 70
const shel_tax_retro_amt            = 71
const shel_tax_retro_verif          = 72
const shel_tax_prosp_amt            = 73
const shel_tax_prosp_verif          = 74
const shel_room_retro_amt           = 75
const shel_room_retro_verif         = 76
const shel_room_prosp_amt           = 77
const shel_room_prosp_verif         = 78
const shel_garage_retro_amt         = 79
const shel_garage_retro_verif       = 80
const shel_garage_prosp_amt         = 81
const shel_garage_prosp_verif       = 82
const shel_subsidy_retro_amt        = 83
const shel_subsidy_retro_verif      = 84
const shel_subsidy_prosp_amt        = 85
const shel_subsidy_prosp_verif      = 86
const shel_notes                    = 87
const shel_verif_checkbox           = 88
const shel_verif_added              = 89
const shel_verif_time               = 90

const memb_notes                    = 91

const new_last_name         = 0
const new_first_name        = 1
const new_mid_initial       = 2
const new_suffix            = 3
const new_full_name         = 4
const new_dob               = 5
const new_rel_to_applicant  = 6
const new_ma_request        = 7
const new_fs_request        = 8
const new_grh_request       = 9
const new_memb_moved_in     = 10
const new_memb_moved_out    = 11
const new_memb_notes        = 12

const ma_request_client     = 0
const ma_request_notes      = 10

Dim NEW_MA_REQUEST_ARRAY()
ReDim NEW_MA_REQUEST_ARRAY(ma_request_notes, 0)


const earned_client         = 0
const earned_type           = 1
const earned_source         = 2
const earned_change_date    = 3
const earned_amount         = 4
const earned_freq           = 5
const earned_hours          = 6
const earned_prog_list      = 7
const earned_start_date     = 8
const earned_seasonal       = 9

const earned_notes          = 11

const unearned_client       = 0
const unearned_type         = 1
const unearned_source       = 2
const unearned_change_date  = 3
const unearned_amount       = 4
const unearned_freq         = 5
Const unearned_prog_list    = 6
const unearned_start_date   = 7
const unearned_notes        = 10

const asset_client          = 0
const asset_type            = 1
const asset_acct_type       = 2
const asset_bank_name       = 3
const asset_year_make_model = 4
const asset_address         = 5
' const asset_
' const asset_
' const asset_
' const asset_
const asset_prog_list       = 9
const asset_notes           = 10

const cs_payer              = 0
const cs_amount             = 1
const cs_current            = 2
const cs_notes              = 10


Dim NEW_EARNED_ARRAY
Dim NEW_UNEARNED_ARRAY
Dim NEW_CHILD_SUPPORT_ARRAY
Dim NEW_ASSET_ARRAY
ReDim NEW_EARNED_ARRAY(earned_notes, 0)
ReDim NEW_UNEARNED_ARRAY(unearned_notes, 0)
ReDim NEW_CHILD_SUPPORT_ARRAY(cs_notes, 0)
ReDim NEW_ASSET_ARRAY(asset_notes, 0)



unea_type_list = "Type or Select"
unea_type_list = unea_type_list+chr(9)+"01 - RSDI, Disa"
unea_type_list = unea_type_list+chr(9)+"02 - RSDI, No Disa"
unea_type_list = unea_type_list+chr(9)+"03 - SSI"
unea_type_list = unea_type_list+chr(9)+"06 - Non-MN PA"
unea_type_list = unea_type_list+chr(9)+"11 - VA Disability"
unea_type_list = unea_type_list+chr(9)+"12 - VA Pension"
unea_type_list = unea_type_list+chr(9)+"13 - VA Other"
unea_type_list = unea_type_list+chr(9)+"38 - VA Aid & Attendance"
unea_type_list = unea_type_list+chr(9)+"14 - Unemployment Insurance"
unea_type_list = unea_type_list+chr(9)+"15 - Worker's Comp"
unea_type_list = unea_type_list+chr(9)+"16 - Railroad Retirement"
unea_type_list = unea_type_list+chr(9)+"17 - Other Retirement"
unea_type_list = unea_type_list+chr(9)+"18 - Military Enrirlement"
unea_type_list = unea_type_list+chr(9)+"19 - FC Child req FS"
unea_type_list = unea_type_list+chr(9)+"20 - FC Child not req FS"
unea_type_list = unea_type_list+chr(9)+"21 - FC Adult req FS"
unea_type_list = unea_type_list+chr(9)+"22 - FC Adult not req FS"
unea_type_list = unea_type_list+chr(9)+"23 - Dividends"
unea_type_list = unea_type_list+chr(9)+"24 - Interest"
unea_type_list = unea_type_list+chr(9)+"25 - Cnt gifts/prizes"
unea_type_list = unea_type_list+chr(9)+"26 - Strike Benefits"
unea_type_list = unea_type_list+chr(9)+"27 - Contract for Deed"
unea_type_list = unea_type_list+chr(9)+"28 - Illegal Income"
unea_type_list = unea_type_list+chr(9)+"29 - Other Countable"
unea_type_list = unea_type_list+chr(9)+"30 - Infrequent"
unea_type_list = unea_type_list+chr(9)+"31 - Other - FS Only"
unea_type_list = unea_type_list+chr(9)+"08 - Direct Child Support"
unea_type_list = unea_type_list+chr(9)+"35 - Direct Spousal Support"
unea_type_list = unea_type_list+chr(9)+"36 - Disbursed Child Support"
unea_type_list = unea_type_list+chr(9)+"37 - Disbursed Spousal Support"
unea_type_list = unea_type_list+chr(9)+"39 - Disbursed CS Arrears"
unea_type_list = unea_type_list+chr(9)+"40 - Disbursed Spsl Sup Arrears"
unea_type_list = unea_type_list+chr(9)+"43 - Disbursed Excess CS"
unea_type_list = unea_type_list+chr(9)+"44 - MSA - Excess Income for SSI"
unea_type_list = unea_type_list+chr(9)+"47 - Tribal Income"
unea_type_list = unea_type_list+chr(9)+"48 - Trust Income"
unea_type_list = unea_type_list+chr(9)+"49 - Non-Recurring"

account_list = "Select or Type"
account_list = account_list+chr(9)+"Cash"
account_list = account_list+chr(9)+"SV - Savings"
account_list = account_list+chr(9)+"CK - Checking"
account_list = account_list+chr(9)+"CE - Certificate of Deposit"
account_list = account_list+chr(9)+"MM - Money Market"
account_list = account_list+chr(9)+"DC - Debit Card"
account_list = account_list+chr(9)+"KO - Keogh Account"
account_list = account_list+chr(9)+"FT - Fed Thrift Savings Plan"
account_list = account_list+chr(9)+"SL - State & Local Govt"
account_list = account_list+chr(9)+"RA - Employee Ret Annuities"
account_list = account_list+chr(9)+"NP - Non-Profit Emmployee Ret"
account_list = account_list+chr(9)+"IR - Indiv Ret Acct"
account_list = account_list+chr(9)+"RH - Roth IRA"
account_list = account_list+chr(9)+"FR - Ret Plan for Employers"
account_list = account_list+chr(9)+"CT - Corp Ret Trust"
account_list = account_list+chr(9)+"RT - Other Ret Fund"
account_list = account_list+chr(9)+"QT - Qualified Tuition (529)"
account_list = account_list+chr(9)+"CA - Coverdell SV (530)"
account_list = account_list+chr(9)+"OE - Other Educational"
account_list = account_list+chr(9)+"OT - Other"

security_list = "Select or Type"
security_list = security_list+chr(9)+"LI - Life Insurance"
security_list = security_list+chr(9)+"ST - Stocks"
security_list = security_list+chr(9)+"BO - Bonds"
security_list = security_list+chr(9)+"CD - Ctrct for Deed"
security_list = security_list+chr(9)+"MO - Mortgage Note"
security_list = security_list+chr(9)+"AN - Annuity"
security_list = security_list+chr(9)+"OT - Other"

cars_list = "Select or Type"
cars_list = cars_list+chr(9)+"1 - Car"
cars_list = cars_list+chr(9)+"2 - Truck"
cars_list = cars_list+chr(9)+"3 - Van"
cars_list = cars_list+chr(9)+"4 - Camper"
cars_list = cars_list+chr(9)+"5 - Motorcycle"
cars_list = cars_list+chr(9)+"6 - Trailer"
cars_list = cars_list+chr(9)+"7 - Other"

rest_list = "Select or Type"
rest_list = rest_list+chr(9)+"1 - House"
rest_list = rest_list+chr(9)+"2 - Land"
rest_list = rest_list+chr(9)+"3 - Buildings"
rest_list = rest_list+chr(9)+"4 - Mobile Home"
rest_list = rest_list+chr(9)+"5 - Life Estate"
rest_list = rest_list+chr(9)+"6 - Other"

state_list = "Select One..."
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
state_list = state_list+chr(9)+"MN Minnesota"
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
'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 190, "Case number dialog"
  EditBox 70, 5, 65, 15, MAXIS_case_number
  EditBox 70, 25, 20, 15, MAXIS_footer_month
  EditBox 95, 25, 20, 15, MAXIS_footer_year
  EditBox 70, 45, 115, 15, Worker_signature
  CheckBox 15, 70, 135, 10, "Check here if this is an exempt (*) IR?", paperless_checkbox
  ButtonGroup ButtonPressed
    OkButton 85, 170, 50, 15
    CancelButton 140, 170, 50, 15
  Text 20, 10, 45, 10, "Case number:"
  Text 20, 30, 45, 10, "Footer Month:"
  Text 125, 30, 25, 10, "mm/yy"
  Text 10, 50, 60, 10, "Worker Signature"
  GroupBox 20, 85, 155, 75, "Exempt IR checkbox warning:"
  Text 25, 100, 145, 25, "If you select ''Is this an exempt IR'', the case note will only provide detail and information about the HC approval."
  Text 25, 135, 140, 20, " If you are processing a CSR with SNAP, you should NOT check that option."
EndDialog
'Showing the case number dialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

Call back_to_SELF
continue_in_inquiry = ""
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If

'If "paperless" was checked, the script will put a simple case note in and end.
If paperless_checkbox = checked then
    run_from_DAIL = FALSE
    call run_from_GitHub(script_repository &  "dail/paperless-dail.vbs")
End If

Call HH_member_custom_dialog(HH_member_array)

Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen GRH_status, 4, 9, 74
EMReadScreen SNAP_status, 4, 10, 74
EMReadScreen HC_status, 4, 12, 74

GRH_active = FALSE
SNAP_active = FALSE
HC_active = FALSE
show_buttons_on_confirmation_dlg = TRUE

If GRH_status = "ACTV" Then GRH_active = TRUE
If SNAP_status = "ACTV" Then SNAP_active = TRUE
If HC_status = "ACTV" Then HC_active = TRUE

'check to see if there is an adult on MA'
Call navigate_to_MAXIS_screen("STAT", "REVW")

grh_sr = FALSE
snap_sr = FALSE
hc_sr = FALSE
'Read for GRH
grh_sr_mo = ""
grh_sr_yr = ""
EMReadScreen grh_revw_status, 1, 7, 40
grh_sr_yn = "No"
If grh_revw_status <> "_" Then
    EMWriteScreen "X", 5, 35
    transmit
    EMReadScreen sr_month, 2, 9, 26
    EMReadScreen sr_year, 2, 9, 32
    PF3

    If grh_revw_status = "N" or grh_revw_status = "I" Then
        grh_sr_mo = sr_month
        grh_sr_yr = sr_year
    Else
        grh_sr_mo = sr_month
        sr_year = sr_year * 1
        sr_year = sr_year - 1
        grh_sr_yr = right("00" & sr_year, 2)
    End If
    grh_sr_yn= "Yes"
End If

'Read for SNAP
snap_sr_mo = ""
snap_sr_yr = ""
curr_snap_sr_status = ""
EMReadScreen snap_revw_status, 1, 7, 60
snap_sr_yn = "No"
If snap_revw_status <> "_" Then
    EMWriteScreen "X", 5, 58
    transmit
    EMReadScreen sr_month, 2, 9, 26
    EMReadScreen sr_year, 2, 9, 32
    PF3
    If snap_revw_status = "N" or snap_revw_status = "I" Then
        snap_sr_mo = sr_month
        snap_sr_yr = sr_year
    Else
        snap_sr_mo = sr_month
        sr_year = sr_year * 1
        sr_year = sr_year - 1
        snap_sr_yr = right("00" & sr_year, 2)
    End If
    snap_sr_yn= "Yes"
End If

'Read for MA
hc_sr_mo = ""
hc_sr_yr = ""
curr_hc_sr_status = ""
EMReadScreen hc_revw_status, 1, 7, 73
hc_sr_yn = "No"
If hc_revw_status <> "_" Then
    EMWriteScreen "X", 5, 71
    transmit
    EMReadScreen ir_month, 2, 8, 27
    EMReadScreen ir_year, 2, 8, 33
    EMReadScreen ar_month, 2, 8, 71
    EMReadScreen ar_year, 2, 8, 77
    PF3
    If ir_month <> "__" Then
        sr_month = ir_month
        sr_year = ir_year
    End If
    If ar_month <> "__" Then
        sr_month = ar_month
        sr_year = ar_year
    End If
    If hc_revw_status = "N" or hc_revw_status = "I" Then
        hc_sr_mo = sr_month
        hc_sr_yr = sr_year
    Else
        hc_sr_mo = sr_month
        sr_year = sr_year * 1
        sr_year = sr_year - 1
        hc_sr_yr = right("00" & sr_year, 2)
    End If
    hc_sr_yn= "Yes"
End If

Call back_to_SELF

Const the_panel_const               = 0
Const the_memb_const                = 1
Const the_inst_const                = 2
Const array_ref_const               = 3
Const panel_btn_const               = 4
Const show_this_panel               = 5
Const one_per_case_const          = 6
Const multiple_per_case_const       = 7
Const one_per_person_const          = 8
Const multiple_per_person_const     = 9
Const panel_notes_const             = 10

Dim ALL_PANELS_ARRAY()
ReDim ALL_PANELS_ARRAY(panel_notes_const, 0)
Call create_array_of_all_panels(ALL_PANELS_ARRAY)

' For the_panels = 0 to UBound(ALL_PANELS_ARRAY, 2)
'     MsgBox "Panel: " & ALL_PANELS_ARRAY(the_panel_const, the_panels) & "-" & ALL_PANELS_ARRAY(the_memb_const, the_panels) & "-" & ALL_PANELS_ARRAY(the_inst_const, the_panels)
' Next

Call generate_client_list(all_the_clients, "Select or Type")
list_for_array = right(all_the_clients, len(all_the_clients) - 15)
full_hh_list = Split(list_for_array, chr(9))
' MsgBox full_hh_list
' MsgBox all_the_clients
' ref_numb
Call back_to_SELF
Call navigate_to_MAXIS_screen("STAT", "MEMB")

' Dim ALL_CLIENTS_ARRAY()
' ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)

member_counter = 0
Do
    EMReadScreen clt_ref_nbr, 2, 4, 33
    EMReadScreen clt_last_name, 25, 6, 30
    EMReadScreen clt_first_name, 12, 6, 63
    EMReadScreen clt_age, 3, 8, 76

    ReDim Preserve ALL_CLIENTS_ARRAY(memb_notes, member_counter)
    ALL_CLIENTS_ARRAY(memb_ref_numb, member_counter) = clt_ref_nbr
    ALL_CLIENTS_ARRAY(memb_last_name, member_counter) = replace(clt_last_name, "_", "")
    ALL_CLIENTS_ARRAY(memb_first_name, member_counter) = replace(clt_first_name, "_", "")
    ALL_CLIENTS_ARRAY(memb_age, member_counter) = trim(clt_age)

    member_counter = member_counter + 1
    transmit
    EMReadScreen last_memb, 7, 24, 2
Loop until last_memb = "ENTER A"

Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, living_situation_status, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, curr_phone_one, curr_phone_two, curr_phone_three, curr_phone_type_one, curr_phone_type_two, curr_phone_type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

new_memb_counter = 0
back_to_dlg_addr		= 1201
back_to_dlg_ma_income	= 1202
back_to_dlg_ma_asset	= 1203
back_to_dlg_snap		= 1204
back_to_dlg_sig			= 1205
add_another_new_memb_btn= 1206
done_adding_new_memb_btn= 1207
add_memb_btn			= 1208
add_jobs_btn			= 1209
add_unea_btn			= 1210
why_answer_btn			= 1211
next_page_ma_btn		= 1212
add_acct_btn			= 1213
add_secu_btn			= 1214
add_cars_btn			= 1215
add_rest_btn			= 1216
back_to_ma_dlg_1		= 1217
continue_btn			= 1218
back_to_ma_dlg_1		= 1219
back_to_ma_dlg_2		= 1220
finish_ma_questions		= 1221
add_snap_earned_income_btn = 1222
add_snap_unearned_btn	= 1223
add_snap_cs_btn			= 1224
complete_csr_questions	= 1225



show_csr_dlg_q_1 		= TRUE
show_csr_dlg_q_2 		= TRUE
show_csr_dlg_q_4_7 		= TRUE
show_csr_dlg_q_9_12 	= TRUE
show_csr_dlg_q_13 		= TRUE
show_csr_dlg_q_15_19 	= TRUE
show_csr_dlg_sig 		= TRUE
show_confirmation		= TRUE

csr_dlg_q_1_cleared 	= FALSE
csr_dlg_q_2_cleared 	= FALSE
csr_dlg_q_4_7_cleared 	= FALSE
csr_dlg_q_9_12_cleared 	= FALSE
csr_dlg_q_13_cleared 	= FALSE
csr_dlg_q_15_19_cleared = FALSE
csr_dlg_sig_cleared 	= FALSE

first_q_1_round = TURE
first_q_2_round = TRUE
questions_answered = FALSE
details_shown = FALSE
abawd_on_case = FALSE

next_page_ma_btn = 1100
previous_page_btn = 1200
continue_btn = 1300

new_earned_counter = 0
new_unearned_counter = 0
new_asset_counter = 0

Dim NEW_MEMBERS_ARRAY()
ReDim NEW_MEMBERS_ARRAY(new_memb_notes, 0)

Do
	Do
		Do
			Do
				Do
					Do
						Do
							Do
								Do
									show_confirmation = TRUE
									If csr_dlg_q_1_cleared = FALSE Then show_csr_dlg_q_1 = TRUE
									If csr_dlg_q_2_cleared = FALSE Then show_csr_dlg_q_2 = TRUE
									If csr_dlg_q_4_7_cleared = FALSE Then show_csr_dlg_q_4_7 = TRUE
									If csr_dlg_q_9_12_cleared = FALSE Then show_csr_dlg_q_9_12 = TRUE
									If csr_dlg_q_13_cleared = FALSE Then show_csr_dlg_q_13 = TRUE
									If csr_dlg_q_15_19_cleared = FALSE Then show_csr_dlg_q_15_19 = TRUE
									If csr_dlg_sig_cleared = FALSE Then show_csr_dlg_sig = TRUE

									If show_csr_dlg_q_1 = TRUE Then Call csr_dlg_q_1

									If first_q_1_round = TURE Then
										Call gather_pers_detail
										first_q_1_round = FALSE
									End If
								Loop until show_csr_dlg_q_1 = FALSE
								If show_csr_dlg_q_2 = TRUE Then Call csr_dlg_q_2

								If first_q_2_round = TURE Then
									Call count_actives
									first_q_2_round = FALSE
								End If
							Loop until show_csr_dlg_q_2 = FALSE
							If show_csr_dlg_q_4_7 = TRUE Then Call csr_dlg_q_4_7
						Loop until show_csr_dlg_q_4_7 = FALSE
						If show_csr_dlg_q_9_12 = TRUE Then Call csr_dlg_q_9_12
					Loop until show_csr_dlg_q_9_12 = FALSE
					If show_csr_dlg_q_13 = TRUE Then Call csr_dlg_q_13
				Loop until show_csr_dlg_q_13 = FALSE
				If show_csr_dlg_q_15_19 = TRUE Then Call csr_dlg_q_15_19
			Loop until show_csr_dlg_q_15_19 = FALSE
			If show_csr_dlg_sig = TRUE Then Call csr_dlg_sig
		Loop until show_csr_dlg_sig = FALSE
		If show_confirmation = TRUE Then Call confirm_csr_form_dlg
	Loop until confirm_csr_form_information = "YES - This is the information on the CSR Form"
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Logic to reflect on if the form is complete.
dlg_width = 260
dlg_len = 140
ma_side_len = 15
'list of all the questions
form_questions_complete = TRUE
ma_questions_complete = TRUE
snap_questions_complete = TRUE
q_one_complete = TRUE

q_two_complete = TRUE
q_four_complete = TRUE
q_five_complete = TRUE
q_six_complete = TRUE
q_seven_complete = TRUE
q_nine_complete = TRUE
q_ten_complete = TRUE
q_eleven_complete = TRUE
q_twelve_complete = TRUE
q_thirteen_complete = TRUE
q_sixteen_complete = TRUE
q_fifteen_complete = TRUE
q_seventeen_complete = TRUE
q_eightneen_complete = TRUE
q_nineteen_complete = TRUE

hh_memb_change = FALSE
If quest_two_move_in_out = "Did not answer" Then
    form_questions_complete = FALSE
    q_two_complete = FALSE
ElseIf quest_two_move_in_out = "Yes" Then
    For known_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
        If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then hh_memb_change = TRUE
        If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then hh_memb_change = TRUE
    Next
    If hh_memb_change = FALSE And new_hh_memb_not_in_mx_yn = "Yes - add another member" Then hh_memb_change = TRUE
    If hh_memb_change = FALSE Then q_two_complete = FALSE
End If
If q_two_complete = FALSE Then dlg_len = dlg_len + 25

If client_on_csr_form = "Person Information Missing" Then
    form_questions_complete = FALSE
    dlg_len = dlg_len + 15
    q_one_complete = FALSE
End If
If residence_address_match_yn = "RESI Address not Provided" Then
    form_questions_complete = FALSE
    dlg_len = dlg_len + 15
    q_one_complete = FALSE
End If
If q_one_complete = FALSE Then dlg_len = dlg_len + 20

If HC_active = TRUE OR hc_sr_yn = "Yes" Then
    ma_side_len = ma_side_len + 50
    dlg_width = 520
    hc_grp_len = 35
    If apply_for_ma = "Did not answer" Then
        ma_questions_complete = FALSE
        q_four_complete = FALSE
    ElseIf apply_for_ma = "Yes" Then
        If q_4_details_blank_checkbox = checked Then
            q_four_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_four_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_four_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_self_employed = "Did not answer" Then
        ma_questions_complete = FALSE
        q_five_complete = FALSE
    ElseIf ma_self_employed = "Yes" Then
        If q_5_details_blank_checkbox = checked Then
            q_five_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_five_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_five_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_start_working = "Did not answer" Then
        ma_questions_complete = FALSE
        q_six_complete = FALSE
    ElseIf ma_start_working = "Yes" Then
        If q_6_details_blank_checkbox = checked Then
            q_six_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_six_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_six_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_other_income = "Did not answer" Then
        ma_questions_complete = FALSE
        q_seven_complete = FALSE
    ElseIf ma_other_income = "Yes" Then
        If q_7_details_blank_checkbox = checked Then
            q_seven_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_seven_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_seven_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_liquid_assets = "Did not answer" Then
        ma_questions_complete = FALSE
        q_nine_complete = FALSE
    ElseIf ma_liquid_assets = "Yes" Then
        If q_9_details_blank_checkbox = checked Then
            q_nine_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_nine_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_nine_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_security_assets = "Did not answer" Then
        ma_questions_complete = FALSE
        q_ten_complete = FALSE
    ElseIf ma_security_assets = "Yes" Then
        If q_10_details_blank_checkbox = checked Then
            q_ten_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_ten_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_ten_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_vehicle = "Did not answer" Then
        ma_questions_complete = FALSE
        q_eleven_complete = FALSE
    ElseIf ma_vehicle = "Yes" Then
        If q_11_details_blank_checkbox = checked Then
            q_eleven_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_eleven_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_eleven_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_real_assets = "Did not answer" Then
        ma_questions_complete = FALSE
        q_twelve_complete = FALSE
    ElseIf ma_real_assets = "Yes" Then
        If q_12_details_blank_checkbox = checked Then
            q_twelve_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_twelve_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_twelve_complete = FALSE Then hc_grp_len = hc_grp_len + 25

    If ma_other_changes = "Did not answer" Then
        ma_questions_complete = FALSE
        q_thirteen_complete = FALSE
    ElseIf ma_other_changes = "Yes" Then
        If changes_reported_blank_checkbox = checked Then
            q_thirteen_complete = FALSE
            ma_questions_complete = False
        End If
    End If
    If q_thirteen_complete = FALSE Then ma_side_len = ma_side_len + 25
    If q_thirteen_complete = FALSE Then hc_grp_len = hc_grp_len + 25
End If
If ma_questions_complete = FALSE Then form_questions_complete = FALSE

If SNAP_active = TRUE OR snap_sr_yn = "Yes" Then
    snap_grp_len = 35
    dlg_len = dlg_len + 50
    If quest_fifteen_form_answer = "Did not answer" Then
        snap_questions_complete = FALSE
        q_fifteen_complete = FALSE
    ElseIf quest_fifteen_form_answer = "Yes" Then
        q_15_details_blank = TRUE
        If trim(new_rent_or_mortgage_amount) <> "" Then q_15_details_blank = FALSE
        If heat_ac_checkbox = checked Then q_15_details_blank = FALSE
        If electricity_checkbox = checked Then q_15_details_blank = FALSE
        If telephone_checkbox = checked Then q_15_details_blank = FALSE

        If q_15_details_blank = TRUE  Then
            q_fifteen_complete = FALSE
        End If
    End If
    If q_fifteen_complete = FALSE Then dlg_len = dlg_len + 25
    If q_fifteen_complete = FALSE Then snap_grp_len = snap_grp_len + 25

    If quest_sixteen_form_answer = "Did not answer" Then
        snap_questions_complete = FALSE
        q_sixteen_complete = FALSE
    ElseIf quest_sixteen_form_answer = "Yes" Then
        If q_16_details_blank_checkbox = checked Then
            snap_questions_complete = FALSE
            q_sixteen_complete = FALSE
        End If
    End If
    If q_sixteen_complete = FALSE  Then dlg_len = dlg_len + 25
    If q_sixteen_complete = FALSE  Then snap_grp_len = snap_grp_len + 25

    If quest_seventeen_form_answer = "Did not answer" Then
        snap_questions_complete = FALSE
        q_seventeen_complete = FALSE
    ElseIf quest_seventeen_form_answer = "Yes" Then
        If q_17_details_blank_checkbox = checked Then
            snap_questions_complete = FALSE
            q_seventeen_complete = FALSE
        End If
    End If
    If q_seventeen_complete = FALSE Then dlg_len = dlg_len + 25
    If q_seventeen_complete = FALSE Then snap_grp_len = snap_grp_len + 25

    If quest_eighteen_form_answer = "Did not answer" Then
        snap_questions_complete = FALSE
        q_eightneen_complete = FALSE
    ElseIf quest_eighteen_form_answer = "Yes" Then
        If q_18_details_blank_checkbox = checked Then
            snap_questions_complete = FALSE
            q_eightneen_complete = FALSE
        End If
    End If
    If q_eightneen_complete = FALSE Then dlg_len = dlg_len + 25
    If q_eightneen_complete = FALSE Then snap_grp_len = snap_grp_len + 25

	If abawd_on_case = TRUE Then
	    If quest_nineteen_form_answer = "Did not answer" Then
	        snap_questions_complete = FALSE
	        q_nineteen_complete = FALSE
	        dlg_len = dlg_len + 25
	        snap_grp_len = snap_grp_len + 25
	    End If
	End If
End If
If snap_questions_complete = FALSE Then form_questions_complete = FALSE

If client_dated_yn = "No" Then
    form_questions_complete = FALSE
    dlg_len = dlg_len + 20
End If
If client_signed_yn = "No" Then
    form_questions_complete = FALSE
    dlg_len = dlg_len + 20
End If

If ma_side_len > dlg_len THen dlg_len = ma_side_len

Do
    err_msg = ""
    y_pos = 90
    hc_y_pos = 40

	Dialog1 = ""
    BeginDialog Dialog1, 0, 0, dlg_width, dlg_len, "CSR Form Complete"
      Text 5, 5, 250, 10, "We have finished gathering what what entered into the CSR form."
      GroupBox 5, 20, 250, 30, "FORM DETAIL ENTRY COMPLETE"
      ' Text 15, 30, 235, 15, "REVIEW AND UPDATE MAXIS. Now that the form has been reviewed, update STAT panels as needed."
      GroupBox 5, 35, 250, 55, "CSR Form Status"
      If form_questions_complete = FALSE Then Text 20, 45, 150, 10, "This CSR Form appears INCOMPLETE."
      If form_questions_complete = TRUE Then Text 20, 45, 150, 10, "This CSR Form appears complete."
	  Text 20, 60, 225, 25, "*** This identifies only if FORM ITSELF HAS BEEN COMPLETED - this does not apply to the whole review. This does not consider verifications or information known to the agency."
      If q_one_complete = FALSE Then
          Text 10, y_pos, 235, 10, "Q1. Name & Address - Question One is incomplete."
          y_pos = y_pos + 10
          If client_on_csr_form = "Person Information Missing" Then
              Text 20, y_pos, 80, 10, "Name is not provided."
              y_pos = y_pos + 10
          End If
          If residence_address_match_yn = "RESI Address not Provided" Then
              Text 20, y_pos, 120, 10, "Residence Address is not provided."
              y_pos = y_pos + 10
          End If
          y_pos = y_pos + 5
      End If

      If q_two_complete = FALSE Then
          Text 10, y_pos, 235, 10, "Q2. Anyone moved in or out - Question Two is incomplete."
          y_pos = y_pos + 10
          If quest_two_move_in_out = "Did not answer" Then
              Text 20, y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
              y_pos = y_pos + 10
          ElseIf quest_two_move_in_out = "Yes" AND hh_memb_change = FALSE THen
              Text 20, y_pos, 170, 10, "Answerred 'Yes' but no detail about a member moving in or out."
              y_pos = y_pos + 10
          End If
          y_pos = y_pos + 5
      End If
      If client_signed_yn = "No" Then
          Text 10, y_pos, 185, 10, "CSR form has not been signed. SIGNATURE MISSING."
          y_pos = y_pos + 15
      End If
      If client_dated_yn = "No" Then
          Text 10, y_pos, 205, 10, "CSR form has not been dated. SIGNATURE DATE MISSING."
          y_pos = y_pos + 15
      End If

      If SNAP_active = TRUE Then
          GroupBox 5, y_pos, 250, snap_grp_len, "Since this case is active SNAP"
          If snap_questions_complete = FALSE Then Text 20, y_pos + 15, 225, 10, "The SNAP portion of the form is INCOMPLETE."
          If snap_questions_complete = TRUE Then Text 20, y_pos + 15, 225, 10, "The SNAP portion of the form is complete."
          y_pos = y_pos + 35
          If q_fifteen_complete = FALSE Then
              Text 10, y_pos, 235, 10, "Q15. Household Moved - Question Fifteen is incomplete."
              Text 20, y_pos + 10, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
              y_pos = y_pos + 25
          End If
          If q_sixteen_complete = FALSE Then
              Text 10, y_pos, 235, 10, "Q16. Job Change - Question Sixteen is incomplete."
              y_pos = y_pos + 10
              If quest_sixteen_form_answer = "Did not answer" Then
                  Text 20, y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  y_pos = y_pos + 10
              ElseIf quest_sixteen_form_answer = "Yes" AND q_16_details_blank_checkbox = checked Then
                  Text 20, y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  y_pos = y_pos + 10
              End If
              y_pos = y_pos + 5
          End If
          If q_seventeen_complete = FALSE Then
              Text 10, y_pos, 235, 10, "Q17. UNEA Change - Question Seventeen is incomplete."
              y_pos = y_pos + 10
              If quest_seventeen_form_answer = "Did not answer" Then
                  Text 20, y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  y_pos = y_pos + 10
              ElseIf quest_seventeen_form_answer = "Yes" AND q_17_details_blank_checkbox = checked Then
                  Text 20, y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  y_pos = y_pos + 10
              End If
              y_pos = y_pos + 5
          End If
          If q_eightneen_complete = FALSE Then
              Text 10, y_pos, 235, 10, "Q18. Child Support Change - Question Eighteen is incomplete."
              y_pos = y_pos + 10
              If quest_eighteen_form_answer = "Did not answer" Then
                  Text 20, y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  y_pos = y_pos + 10
              ElseIf quest_eighteen_form_answer = "Yes" and q_18_details_blank_checkbox = checked Then
                  Text 20, y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  y_pos = y_pos + 10
              End If
              y_pos = y_pos + 5
          End If
          If q_nineteen_complete = FALSE Then
              Text 10, y_pos, 235, 10, "Q19. Child Support Change - Question Ninetee is incomplete."
              Text 20, y_pos + 10, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
              y_pos = y_pos + 25
          End If
      End If


      If HC_active = TRUE Then
          GroupBox 265, 5, 250, hc_grp_len, "Since this case is active HC"
          If ma_questions_complete = FALSE Then Text 280, 20, 225, 10, "The HC portion of the form is INCOMPLETE."
          If ma_questions_complete = TRUE Then Text 280, 20, 225, 10, "The HC portion of the form is complete."
          If q_four_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q4. Apply for new MA Coverage - Question Four is incomplete."
              hc_y_pos = hc_y_pos + 10
              If apply_for_ma = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_five_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q5. Self-Employed - Question Fve is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_self_employed = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_self_employed = "Yes" AND q_5_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_six_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q6.Working - Question Six is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_start_working = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_start_working = "Yes" AND q_6_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_seven_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q7. Unearned Income - Question Seven is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_other_income = "Did not_answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_other_income = "Yes" AND q_7_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_nine_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q9. Bank Account - Question Nine is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_liquid_assets = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_liquid_assets = "Yes" AND q_9_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_ten_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q10. Securities - Question Ten is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_security_assets = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_security_assets = "Yes" and q_10_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_eleven_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q11.Vehicle - Question Eleven is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_vehicle = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_vehicle = "Yes" AND q_11_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_twelve_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q12. Real Estate - Question Twelve is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_real_assets = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_real_assets = "Yes" and q_12_details_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
          If q_thirteen_complete = FALSE Then
              Text 270, hc_y_pos, 235, 10, "Q13. Changes - Question Thirteen is incomplete."
              hc_y_pos = hc_y_pos + 10
              If ma_other_changes = "Did not answer" Then
                  Text 280, hc_y_pos, 170, 10, "Requires 'Yes' or 'No' and neither were checked."
                  hc_y_pos = hc_y_pos + 10
              ElseIf ma_other_changes = "Yes" and changes_reported_blank_checkbox = checked Then
                  Text 280, hc_y_pos, 170, 10, "Answered 'Yes' but detail was not provided."
                  hc_y_pos = hc_y_pos + 10
              End If
              hc_y_pos = hc_y_pos + 5
          End If
      End If
      CheckBox 10, dlg_len - 40, 200, 10, "Check here if this assessment is WRONG.", functionality_wrong_checkbox
      CheckBox 10, dlg_len - 25, 110, 10, "Check here to have this detail ", export_form_info_to_word_checkbox
      Text 25, dlg_len - 15, 75, 10, "exported to Word."

      ButtonGroup ButtonPressed
        ' PushButton 10, dlg_len - 15, 100, 10, "Export to Word", export_to_word_btn
        ' OkButton dlg_width - 110, dlg_len - 20, 50, 15
        ' PushButton dlg_width - 120, dlg_len - 20, 60, 15, "Panels Updated", panels_updated_btn
		OkButton dlg_width - 120, dlg_len - 20, 60, 15
        CancelButton dlg_width - 55, dlg_len - 20, 50, 15
    EndDialog

    ' MsgBox y_pos & " - y pos"
    ' MsgBox dlg_len & " - dlg len"
    dialog Dialog1
    cancel_confirmation

    ' If ButtonPressed = -1 Then err_msg = "LOOP"

Loop until err_msg = ""

If functionality_wrong_checkbox = checked Then
    If form_questions_complete = FALSE Then form_completion_status = "Incomplete"
    If form_questions_complete = TRUE Then form_completion_status = "Complete"
    If snap_questions_complete = FALSE Then snap_completion_status = "Incomplete"
    If snap_questions_complete = TRUE Then snap_completion_status = "Complete"
    If SNAP_active = FALSE Then snap_completion_status = "SNAP not relevant"
    If ma_questions_complete = FALSE Then hc_completion_status = "Incomplete"
    If ma_questions_complete = TRUE Then hc_completion_status = "Complete"
    If HC_active = FALSE Then hc_completion_status = "HC not relevant"

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 386, 130, "Form Complettion Incorrect"
      DropListBox 105, 40, 85, 45, " "+chr(9)+"Complete"+chr(9)+"Incomplete"+chr(9)+"Unsure", form_completion_status
      DropListBox 260, 40, 85, 45, " "+chr(9)+"Complete"+chr(9)+"Incomplete"+chr(9)+"Unsure"+chr(9)+"SNAP not relevant", snap_completion_status
      DropListBox 260, 60, 85, 45, " "+chr(9)+"Complete"+chr(9)+"Incomplete"+chr(9)+"Unsure"+chr(9)+"HC not relevant", hc_completion_status
      EditBox 5, 90, 375, 15, functionality_issue_notes
      ButtonGroup ButtonPressed
        OkButton 330, 110, 50, 15
      Text 5, 10, 155, 10, "Indicate the completion status of the form itself."
      Text 10, 20, 180, 10, "This is not if the SR process is complete - just the form."
      Text 20, 45, 80, 10, "The form as a whole is "
      Text 205, 45, 50, 10, "SNAP Portion"
      Text 215, 65, 35, 10, "HC Portion"
      Text 5, 80, 135, 10, "Add detail about the error in functionality"
    EndDialog

    Do
        dialog Dialog1
        cancel_confirmation

        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    email_msg = "The script determined completion functionality did not function as expected for case: " & MAXIS_case_number & vbCr & vbCr
    email_msg = email_msg & "Script Found:" & vbCr
    If form_questions_complete = TRUE Then email_msg = email_msg & "Form itself is COMPLETE" & vbCr
    If form_questions_complete = FALSE Then email_msg = email_msg & "Form itself is INCOMPLETE" & vbCr
    If snap_questions_complete = TRUE Then email_msg = email_msg & "SNAP portion of the form is COMPLETE" & vbCr
    If snap_questions_complete = FALSE Then email_msg = email_msg & "SNAP portion of the form is INCOMPLETE" & vbCr
    If SNAP_active = FALSE Then email_msg = email_msg & "SNAP is INACTIVE" & vbCr
    If ma_questions_complete = TRUE Then email_msg = email_msg & "HC portion of the form is COMPLETE" & vbCr
    If ma_questions_complete = FALSE Then email_msg = email_msg & "HC portion of the form is INCOMPLETE" & vbCr
    If HC_active = FALSE Then email_msg = email_msg & "HC is INACTIVE" & vbCr & vbCR

    email_msg = email_msg & "-------------------------------------" & vbCr

    email_msg = email_msg & "Question 1: Name and Address" & vbCr
    If client_on_csr_form = "Person Information Missing" Then
        email_msg = email_msg & "   Person/name missing from the form" & vbCr
    Else
        email_msg = email_msg & "   The name listed on the form: " & client_on_csr_form & vbCr
    End If
    If residence_address_match_yn = "RESI Address not Provided" Then
        email_msg = email_msg & "   Address information not provided on the form" & vbCr
    Else
        email_msg = email_msg & "   Residence address provided on the form and it " & residence_address_match_yn & vbCr
    End If
    email_msg = email_msg & "Question 2: Anyone move in our out? Answer - " & quest_two_move_in_out & vbCr
    email_msg = email_msg & "   Household Members Change Indicated - " & hh_memb_change & vbCr
    email_msg = email_msg & "Question 4: Anyone else applying for MA? Answer - " & apply_for_ma & vbCr
    If q_4_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 4 were left blank." & vbCr
    email_msg = email_msg & "Question 5: Anyone Self-Employed? Answer - " & ma_self_employed & vbCr
    If q_5_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 5 were left blank." & vbCr
    email_msg = email_msg & "Question 6: Anyone working? Answer - " & ma_start_working & vbCr
    If q_6_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 6 were left blank." & vbCr
    email_msg = email_msg & "Question 7: Anyone have other income? Answer - " & ma_other_income & vbCr
    If q_7_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 7 were left blank." & vbCr
    email_msg = email_msg & "Question 9: Anyone have liquid assets? Answer - " & ma_liquid_assets & vbCr
    If q_9_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 9 were left blank." & vbCr
    email_msg = email_msg & "Question 10: Anyone have security assets? Answer - " & ma_security_assets & vbCr
    If q_10_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 10 were left blank." & vbCr
    email_msg = email_msg & "Question 11: Anyone have a vehicle? Answer - " & ma_vehicle & vbCr
    If q_11_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 11 were left blank." & vbCr
    email_msg = email_msg & "Question 12: Anyone have real estate? Answer - " & ma_real_assets & vbCr
    If q_12_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 12 were left blank." & vbCr
    email_msg = email_msg & "Question 13: Other changes to report? Answer - " & ma_other_changes & vbCr
    If changes_reported_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 13 were left blank." & vbCr

    email_msg = email_msg & "Question 15: Has the houshold moved? Answer - " & quest_fifteen_form_answer & vbCr
    If q_15_details_blank = TRUE Then email_msg = email_msg & "   Details of Question 15 were left blank (no mortgage/rent or utilities information provided)." & vbCr
    email_msg = email_msg & "Question 16: Has there been a change in Earned Income? Answer - " & quest_sixteen_form_answer & vbCr
    If q_16_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 16 were left blank." & vbCr
    email_msg = email_msg & "Question 17: Has there been a change in Unearned Income? Answer - " & quest_seventeen_form_answer & vbCr
    If q_17_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 17 were left blank." & vbCr
    email_msg = email_msg & "Question 18: Has there been a change in Child Support? Answer - " &quest_eighteen_form_answer  & vbCr
    If q_18_details_blank_checkbox = checked Then email_msg = email_msg & "   Details of Question 19 were left blank." & vbCr
    email_msg = email_msg & "Question 19: Did you work 20 hours per week? Answer - " & quest_nineteen_form_answer & vbCr

    email_msg = email_msg & "CSR Form was Signed - " & client_signed_yn & vbCr
    email_msg = email_msg & "CSR Form was Dated - " & client_dated_yn & vbCr & vbCr
    email_msg = email_msg & "-------------------------------------" & vbCr

    email_msg = email_msg & "After manual review, the following is what appears to be the actual completion status:" & vbCr
    email_msg = email_msg & "The actual form is: " & form_completion_status & vbCr
    email_msg = email_msg & "SNAP Portion of the form is: " & snap_completion_status & vbCr
    email_msg = email_msg & "HC Portion of the form is: " & hc_completion_status & vbCr & vbCr
    email_msg = email_msg & "Notes: " & functionality_issue_notes & vbCr & vbCR
    email_msg = email_msg & worker_signature & vbCr

    Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", "NOTES - CSR Completion Functionality Issues", email_msg, "", TRUE)

    If form_completion_status = "Complete" Then form_questions_complete = TRUE
    If form_completion_status = "Incomplete" Then form_questions_complete = FALSE

    If snap_completion_status = "Complete" Then snap_questions_complete = TRUE
    If snap_completion_status = "Incomplete" Then snap_questions_complete = FALSE
    If snap_completion_status = "SNAP not relevant" Then SNAP_active = FALSE

    If hc_completion_status = "Complete" Then ma_questions_complete = TRUE
    If hc_completion_status = "Incomplete" Then ma_questions_complete = FALSE
    If hc_completion_status = "HC note relevant" Then HC_active = FALSE

End If

show_buttons_on_confirmation_dlg = FALSE
Do
	Do
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 281, 135, "Update MAXIS NOW"
		  ButtonGroup ButtonPressed
		    PushButton 135, 90, 140, 15, "Show CSR Details", show_confirmation_btn
		    PushButton 35, 115, 180, 15, "MAXIS Panels for CSR have been Updated", all_panels_updated_btn
		  Text 10, 10, 265, 10, "Now that we have reviewed the CSR Form, update the MAXIS panels for " & MAXIS_footer_month & "/" & MAXIS_footer_year & "."
		  Text 25, 25, 240, 20, "*** The script will read MAXIS panels after you have completed the update and pressed the button to indicate the updates are complete. "
		  Text 25, 50, 240, 20, "The next step in the script will allow you add notes about the details in MAXIS and indicate verifications needed."
		  Text 5, 75, 270, 10, "The CASE:NOTEs will only be entered after completion of the notes and information."
		  Text 5, 95, 130, 10, "To review details from the CSR FORM:"
		EndDialog

		dialog Dialog1

		If ButtonPressed = show_confirmation_btn Then Call confirm_csr_form_dlg
	Loop until ButtonPressed = all_panels_updated_btn

	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call navigate_to_MAXIS_screen("STAT", "MEMB")
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
    transmit

    EMReadScreen clt_id_verif, 2, 9, 68
    EMReadScreen clt_ssn, 11, 7, 42
    EMReadScreen clt_ssn_verif, 1, 7, 68
    EMReadScreen clt_dob, 10, 8, 42
    EMReadScreen clt_dob_verif, 2, 8, 68
    EMReadScreen clt_gender, 1, 9, 42

    EMReadScreen clt_rel_to_applct, 2, 10, 42
    EMReadScreen clt_spkn_lang, 20, 12, 42
    EMReadScreen clt_wrt_lang, 29, 13, 42
    EMReadScreen clt_interp_need, 1, 14, 68
    EMReadScreen clt_alias, 1, 15, 42
    EMReadScreen clt_ethncty, 1, 16, 68
    EMReadScreen clt_race, 30, 17, 42

    If clt_id_verif = "BC" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "BC - Birth Certificate"
    If clt_id_verif = "RE" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "RE - Religious Record"
    If clt_id_verif = "DL" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "DL - Drivers Lic/St ID"
    If clt_id_verif = "DV" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "DV - Divorce Decree"
    If clt_id_verif = "AL" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "AL - Alien Card"
    If clt_id_verif = "AD" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "AD - Arrival/Depart"
    If clt_id_verif = "DR" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "DR - Doctor Stmt"
    If clt_id_verif = "PV" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "PV = Passport/Visa"
    If clt_id_verif = "OT" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "OT - Other"
    If clt_id_verif = "NO" Then ALL_CLIENTS_ARRAY(memb_id_verif, case_memb) = "NO - No Ver Prvd"
    ALL_CLIENTS_ARRAY(memb_soc_sec_numb, case_memb) = replace(clt_ssn, " ", "-")
    If clt_ssn_verif = "A" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "A - SSN Applied For"
    If clt_ssn_verif = "P" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "P - SSN Prvd, Verif Pending"
    If clt_ssn_verif = "N" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "N - SSN Not Prvd"
    If clt_ssn_verif = "V" Then ALL_CLIENTS_ARRAY(memb_ssn_verif, case_memb) = "V - System Verified"
    ALL_CLIENTS_ARRAY(memb_dob, case_memb) = replace(clt_dob, " ", "/")
    If clt_dob_verif = "BC" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "BC - Birth Certificate"
    If clt_dob_verif = "RE" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "RE - Religious Record"
    If clt_dob_verif = "DL" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "DL - Drivers Lic/St ID"
    If clt_dob_verif = "DV" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "DV - Divorce Decree"
    If clt_dob_verif = "AL" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "AL - Alien Card"
    If clt_dob_verif = "DR" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "DR - Doctor Stmt"
    If clt_dob_verif = "PV" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "PV = Passport/Visa"
    If clt_dob_verif = "OT" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "OT - Other"
    If clt_dob_verif = "NO" Then ALL_CLIENTS_ARRAY(memb_dob_verif, case_memb) = "NO - No Ver Prvd"
    If clt_gender = "F" Then ALL_CLIENTS_ARRAY(memb_gender, case_memb) = "Female"
    If clt_gender = "M" Then ALL_CLIENTS_ARRAY(memb_gender, case_memb) = "Male"
    If clt_rel_to_applct = "01" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "01 - Applicant"
    If clt_rel_to_applct = "02" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "02 - Spouse"
    If clt_rel_to_applct = "03" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "03 - Child"
    If clt_rel_to_applct = "04" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "04 - Parent"
    If clt_rel_to_applct = "05" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "05 - Sibling"
    If clt_rel_to_applct = "06" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "06 - Step Sibling"
    If clt_rel_to_applct = "08" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "08 - Step Child"
    If clt_rel_to_applct = "09" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "09 - Step Parent"
    If clt_rel_to_applct = "10" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "10 - Aunt"
    If clt_rel_to_applct = "11" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "11 - Uncle"
    If clt_rel_to_applct = "12" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "12 - Niece"
    If clt_rel_to_applct = "13" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "13 - Nephew"
    If clt_rel_to_applct = "14" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "14 - Cousin"
    If clt_rel_to_applct = "15" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "15 - Grandparent"
    If clt_rel_to_applct = "16" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "16 - Grandchild"
    If clt_rel_to_applct = "17" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "17 - Other Relative"
    If clt_rel_to_applct = "18" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "18 - Legal Guardian"
    If clt_rel_to_applct = "24" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "24 - Not Related"
    If clt_rel_to_applct = "25" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "25 - Live-In Attendant"
    If clt_rel_to_applct = "27" Then ALL_CLIENTS_ARRAY(memb_rel_to_applct, case_memb) = "27 - Unknown"

    clt_spkn_lang = replace(clt_spkn_lang, "_", "")
    clt_spkn_lang = replace(clt_spkn_lang, "  ", " - ")
    ALL_CLIENTS_ARRAY(memb_spoken_language, case_memb) = trim(clt_spkn_lang)
    clt_wrt_lang = replace(clt_wrt_lang, "_", "")
    clt_wrt_lang = replace(clt_wrt_lang, "  ", " - ")
    clt_wrt_lang = replace(clt_wrt_lang, "(HRF)", "")
    ALL_CLIENTS_ARRAY(memb_written_language, case_memb) = trim(clt_wrt_lang)

    ALL_CLIENTS_ARRAY(memb_interpreter, case_memb) = clt_interp_need
    ALL_CLIENTS_ARRAY(memb_alias, case_memb) = clt_alias
    ALL_CLIENTS_ARRAY(memb_ethnicity, case_memb) = clt_ethncty
    ALL_CLIENTS_ARRAY(memb_race, case_memb) = trim(clt_race)
Next

Call navigate_to_MAXIS_screen("STAT", "MEMI")
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
    transmit

    EMReadScreen clt_mar_status, 1, 7, 40
    EMReadScreen clt_spouse, 2, 9, 49

    EMReadScreen clt_desg_spouse_yn, 1, 7, 71
    EMReadScreen clt_marriage_date, 8, 8, 40
    EMReadScreen clt_marriage_date_verif, 8, 8, 71

    EMReadScreen clt_citizen, 1, 11, 49
    EMReadScreen clt_cit_verif, 2, 11, 78
    EMReadScreen clt_last_grade, 2, 10, 49
    EMReadScreen clt_in_MN_12_mo, 1, 14, 49
    EMReadScreen clt_resi_verif, 1, 14, 78
    EMReadScreen clt_MN_entry_date, 8, 15, 49
    EMReadScreen clt_former_state, 2, 15, 78
    EMReadScreen clt_other_st_FS_end, 8, 13, 49

    If clt_mar_status = "N" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Never married"
    If clt_mar_status = "M" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Married, Living with Spouse"
    If clt_mar_status = "S" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Married Living Apart"
    If clt_mar_status = "L" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Legally Separated"
    If clt_mar_status = "D" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Divorced"
    If clt_mar_status = "W" Then ALL_CLIENTS_ARRAY(memi_marriage_status, case_memb) = "Widowed"
    ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) = replace(clt_spouse, "_", "")
    If ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) <> "" Then
        For all_the_people = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
            If ALL_CLIENTS_ARRAY(memb_ref_nbr, all_the_people) = ALL_CLIENTS_ARRAY(memi_spouse_ref, case_memb) Then
                ALL_CLIENTS_ARRAY(memi_spouse_name, case_memb) = ALL_CLIENTS_ARRAY(memb_first_name, all_the_people) & " " & ALL_CLIENTS_ARRAY(memb_last_name, all_the_people)
            End If
        Next
    End If
    ALL_CLIENTS_ARRAY(memi_designated_spouse, case_memb) = replace(clt_desg_spouse_yn, "_", "")
    ALL_CLIENTS_ARRAY(memi_marriage_date, case_memb) = replace(clt_marriage_date, " ", "/")
    ALL_CLIENTS_ARRAY(memi_marriage_verif, case_memb) = replace(clt_marriage_date_verif, " ", "/")
    ALL_CLIENTS_ARRAY(memi_citizen, case_memb) = clt_citizen
    If clt_cit_verif = "BC" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "BC - Birth Certificate"
    If clt_cit_verif = "RE" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "RE - Religious Record"
    If clt_cit_verif = "NP" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "NP - Naturalization Papers"
    If clt_cit_verif = "IM" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "IM - Immigration Document"
    If clt_cit_verif = "PV" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "PV - Passport/Visa"
    If clt_cit_verif = "OT" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "OT - Other Document"
    If clt_cit_verif = "NO" Then ALL_CLIENTS_ARRAY(memi_citizen_verif, case_memb) = "NO - No Ver prvd"

    If clt_last_grade = "00" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Pre 1st Grd"
    If clt_last_grade = "01" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 1"
    If clt_last_grade = "02" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 2"
    If clt_last_grade = "03" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 3"
    If clt_last_grade = "04" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 4"
    If clt_last_grade = "05" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 5"
    If clt_last_grade = "06" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 6"
    If clt_last_grade = "07" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 7"
    If clt_last_grade = "08" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 8"
    If clt_last_grade = "09" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 9"
    If clt_last_grade = "10" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 10"
    If clt_last_grade = "11" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grade 11"
    If clt_last_grade = "12" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "HS Diploma or GED"
    If clt_last_grade = "13" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Some Post Sec Ed"
    If clt_last_grade = "14" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "High Schl Plus Cert"
    If clt_last_grade = "15" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Four Yr Degree"
    If clt_last_grade = "16" Then ALL_CLIENTS_ARRAY(memi_last_grade, case_memb) = "Grad Degree"

    ALL_CLIENTS_ARRAY(memi_in_MN_less_12_mo, case_memb) = clt_in_MN_12_mo
    If ALL_CLIENTS_ARRAY(memi_in_MN_less_12_mo, case_memb) = "__/__/__" Then ALL_CLIENTS_ARRAY(memi_in_MN_less_12_mo, case_memb) = ""
    IF clt_resi_verif = "1" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "1 - Rent Receipt"
    IF clt_resi_verif = "2" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "2 - Landlord's Stmt"
    IF clt_resi_verif = "3" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "3 - Utility Bill"
    IF clt_resi_verif = "4" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "4 - Other"
    IF clt_resi_verif = "N" Then ALL_CLIENTS_ARRAY(memi_resi_verif, case_memb) = "N - Ver Not Prvd"
    ALL_CLIENTS_ARRAY(memi_MN_entry_date, case_memb) = replace(clt_MN_entry_date, " ", "/")
    If ALL_CLIENTS_ARRAY(memi_MN_entry_date, case_memb) = "__/__/__" Then ALL_CLIENTS_ARRAY(memi_MN_entry_date, case_memb) = ""
    ALL_CLIENTS_ARRAY(memi_former_state, case_memb) = replace(clt_former_state, "_", "")
    ALL_CLIENTS_ARRAY(memi_other_FS_end, case_memb) = replace(clt_other_st_FS_end, " ", "/")
    If ALL_CLIENTS_ARRAY(memi_other_FS_end, case_memb) = "__/__/__" Then ALL_CLIENTS_ARRAY(memi_other_FS_end, case_memb) = ""

Next

Call navigate_to_MAXIS_screen("STAT", "WREG")
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
    transmit

    EMReadScreen panel_exists, 14, 24, 13
    If panel_exists <> "DOES NOT EXIST" Then
        call access_WREG_panel("READ", notes_on_wreg, ALL_CLIENTS_ARRAY(wreg_pwe, case_memb), ALL_CLIENTS_ARRAY(wreg_status, case_memb), ALL_CLIENTS_ARRAY(wreg_defer_fset, case_memb), ALL_CLIENTS_ARRAY(wreg_fset_orient_date, case_memb), ALL_CLIENTS_ARRAY(wreg_sanc_begin_date, case_memb), ALL_CLIENTS_ARRAY(wreg_sanc_count, case_memb), ALL_CLIENTS_ARRAY(wreg_sanc_reasons, case_memb), ALL_CLIENTS_ARRAY(wreg_abawd_status, case_memb), ALL_CLIENTS_ARRAY(wreg_banekd_months, case_memb), ALL_CLIENTS_ARRAY(wreg_GA_basis, case_memb), ALL_CLIENTS_ARRAY(wreg_GA_coop, case_memb),  ALL_CLIENTS_ARRAY(wreg_numb_ABAWD_months, case_memb),  ALL_CLIENTS_ARRAY(wreg_ABAWD_months_list, case_memb),  ALL_CLIENTS_ARRAY(wreg_numb_second_set_months, case_memb),  ALL_CLIENTS_ARRAY(wreg_second_set_months_list, case_memb))
    End If
Next

Call navigate_to_MAXIS_screen("STAT", "SHEL")
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
    transmit

    EMReadScreen panel_exists, 14, 24, 13
    If panel_exists <> "DOES NOT EXIST" Then
        call access_SHEL_panel("READ", ALL_CLIENTS_ARRAY(shel_hud_sub_yn, case_memb), ALL_CLIENTS_ARRAY(shel_shared_yn, case_memb), ALL_CLIENTS_ARRAY(shel_paid_to, case_memb), ALL_CLIENTS_ARRAY(shel_rent_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_rent_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_rent_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_rent_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_lot_rent_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_lot_rent_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_mortgage_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_mortgage_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_mortgage_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_mortgage_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_insurance_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_insurance_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_insurance_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_insurance_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_tax_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_tax_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_tax_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_tax_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_room_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_room_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_room_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_room_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_garage_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_garage_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_garage_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_garage_prosp_verif, case_memb), ALL_CLIENTS_ARRAY(shel_subsidy_retro_amt, case_memb), ALL_CLIENTS_ARRAY(shel_subsidy_retro_verif, case_memb), ALL_CLIENTS_ARRAY(shel_subsidy_prosp_amt, case_memb), ALL_CLIENTS_ARRAY(shel_subsidy_prosp_verif, case_memb))
    End If
Next

Call access_HEST_panel("READ", HEST_persons_paying, HEST_fs_choice_date, HEST_initial_month_actual_expense, HEST_retro_heat_air, HEST_retro_heat_air_units, HEST_retro_heat_air_amount, HEST_retro_electric, HEST_retro_electric_units, HEST_retro_electric_amount, HEST_retro_phone, HEST_retro_phone_units, HEST_retro_phone_amount, HEST_prosp_heat_air, HEST_prosp_heat_air_units, HEST_prosp_heat_air_amount, HEST_prosp_electric, HEST_prosp_electric_units, HEST_prosp_electric_amount, HEST_prosp_phone, HEST_prosp_phone_units, HEST_prosp_phone_amount, HEST_total_expense)

Call Navigate_to_MAXIS_screen("STAT", "FACI")
faci_panels = 0
For case_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
    EMWriteScreen ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb), 20, 76
    EmWriteScreen "01", 20, 79
    transmit

    Do
        EMReadScreen last_faci_panel, 7, 24, 2
        If last_faci_panel <> "ENTER A" Then
            ReDim Preserve FACILITIES_ARRAY(faci_notes, faci_panels)
            FACILITIES_ARRAY(faci_ref_numb, faci_panels) = ALL_CLIENTS_ARRAY(memb_ref_numb, case_memb)
            FACILITIES_ARRAY(faci_member, faci_panels) = ALL_CLIENTS_ARRAY(memb_first_name, case_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, case_memb)
            EMReadScreen this_instance, 1,2, 73
            FACILITIES_ARRAY(faci_instance, faci_panels) = "0" & this_instance

            Call access_FACI_panel("READ", notes_on_faci, FACILITIES_ARRAY(faci_name, faci_panels), FACILITIES_ARRAY(faci_vendor_number, faci_panels), FACILITIES_ARRAY(faci_type, faci_panels), FACILITIES_ARRAY(faci_FS_elig, faci_panels), FACILITIES_ARRAY(faci_FS_type, faci_panels), FACILITIES_ARRAY(faci_waiver_type, faci_panels), FACILITIES_ARRAY(faci_ltc_inelig_reason, faci_panels), FACILITIES_ARRAY(faci_inelig_begin_date, faci_panels), FACILITIES_ARRAY(faci_inelig_end_date, faci_panels), FACILITIES_ARRAY(faci_anticipated_out_date, faci_panels), FACILITIES_ARRAY(faci_GRH_plan_required, faci_panels), FACILITIES_ARRAY(faci_GRH_plan_verif, faci_panels), FACILITIES_ARRAY(faci_cty_app_place, faci_panels), FACILITIES_ARRAY(faci_approval_cty_name, faci_panels), FACILITIES_ARRAY(faci_GRH_DOC_amount, faci_panels), FACILITIES_ARRAY(faci_GRH_postpay, faci_panels), FACILITIES_ARRAY(faci_stay_one_rate, faci_panels), FACILITIES_ARRAY(faci_stay_one_date_in, faci_panels), FACILITIES_ARRAY(faci_stay_one_date_out, faci_panels), FACILITIES_ARRAY(faci_stay_two_rate, faci_panels), FACILITIES_ARRAY(faci_stay_two_date_in, faci_panels), FACILITIES_ARRAY(faci_stay_two_date_out, faci_panels), FACILITIES_ARRAY(faci_stay_three_rate, faci_panels), FACILITIES_ARRAY(faci_stay_three_date_in, faci_panels), FACILITIES_ARRAY(faci_stay_three_date_out, faci_panels), FACILITIES_ARRAY(faci_stay_four_rate, faci_panels), FACILITIES_ARRAY(faci_stay_four_date_in, faci_panels), FACILITIES_ARRAY(faci_stay_four_date_out, faci_panels), FACILITIES_ARRAY(faci_stay_five_rate, faci_panels), FACILITIES_ARRAY(faci_stay_five_date_in, faci_panels), FACILITIES_ARRAY(faci_stay_five_date_out, faci_panels))

            faci_panels = faci_panels + 1
            transmit
        End If
    Loop until last_faci_panel = "ENTER A"
Next

Call navigate_to_MAXIS_screen("STAT", "REVW")

panel_grh_status = ""
EMReadScreen grh_revw_status, 1, 7, 40

EMWriteScreen "X", 5, 35
transmit
EMReadScreen panel_grh_sr_month, 2, 9, 26
EMReadScreen panel_grh_sr_year, 2, 9, 32
EMReadScreen panel_grh_er_month, 2, 9, 64
EMReadScreen panel_grh_er_year, 2, 9, 70
PF3

If grh_revw_status = "I" Then panel_grh_status = "I - Incomplete"
If grh_revw_status = "U" Then panel_grh_status = "U - Complete and Updt Req"
If grh_revw_status = "N" Then panel_grh_status = "N - Not Rcvd"
If grh_revw_status = "A" Then panel_grh_status = "A - Approved"
If grh_revw_status = "O" Then panel_grh_status = "O - Override Autoclose"
If grh_revw_status = "T" Then panel_grh_status = "T - Terminated"
If grh_revw_status = "D" Then panel_grh_status = "D - Denied"
If grh_revw_status = "_" Then grh_revw_status = "Not Due"

panel_snap_status = ""
EMReadScreen snap_revw_status, 1, 7, 60

EMWriteScreen "X", 5, 58
transmit
EMReadScreen panel_snap_sr_month, 2, 9, 26
EMReadScreen panel_snap_sr_year, 2, 9, 32
EMReadScreen panel_snap_er_month, 2, 9, 64
EMReadScreen panel_snap_er_year, 2, 9, 70
PF3

If snap_revw_status = "I" Then panel_snap_status = "I - Incomplete"
If snap_revw_status = "U" Then panel_snap_status = "U - Complete and Updt Req"
If snap_revw_status = "N" Then panel_snap_status = "N - Not Rcvd"
If snap_revw_status = "A" Then panel_snap_status = "A - Approved"
If snap_revw_status = "O" Then panel_snap_status = "O - Override Autoclose"
If snap_revw_status = "T" Then panel_snap_status = "T - Terminated"
If snap_revw_status = "D" Then panel_snap_status = "D - Denied"
If snap_revw_status = "_" Then panel_snap_status = "Not Due"

panel_hc_status = ""
EMReadScreen hc_revw_status, 1, 7, 73

EMWriteScreen "X", 5, 71
transmit
EMReadScreen ir_month, 2, 8, 27
EMReadScreen ir_year, 2, 8, 33
EMReadScreen ar_month, 2, 8, 71
EMReadScreen ar_year, 2, 8, 77
EMReadScreen panel_hc_er_month, 2, 9, 27
EMReadScreen panel_hc_er_year, 2, 9, 33
PF3
If ir_month <> "__" Then
    panel_hc_sr_month = ir_month
    panel_hc_sr_year = ir_year
End If
If ar_month <> "__" Then
    panel_hc_sr_month = ar_month
    panel_hc_sr_year = ar_year
End If

If hc_revw_status = "I" Then panel_hc_status = "I - Incomplete"
If hc_revw_status = "U" Then panel_hc_status = "U - Complete and Updt Req"
If hc_revw_status = "N" Then panel_hc_status = "N - Not Rcvd"
If hc_revw_status = "A" Then panel_hc_status = "A - Approved"
If hc_revw_status = "O" Then panel_hc_status = "O - Override Autoclose"
If hc_revw_status = "T" Then panel_hc_status = "T - Terminated"
If hc_revw_status = "D" Then panel_hc_status = "D - Denied"
If hc_revw_status = "_" Then hc_revw_status = "Not Due"
' MsgBox "SNAP SR - " & panel_snap_sr_month & "/" & panel_snap_sr_year & vbNewLine & "SNAP ER - " & panel_snap_er_month & "/" & panel_snap_er_year

Call navigate_to_MAXIS_screen("STAT", "MONT")
EMReadScreen mont_versions, 1, 2, 73
mont_detail = ""
If mont_versions = "0" Then
    mont_detail = "This case does not have monthly reporting."
Else
    EMReadScreen mont_date_received, 8, 6, 39
    EMReadScreen mont_cash_status, 1, 11, 43
    EMReadScreen mont_snap_status, 1, 11, 53
    EMReadScreen mont_hc_status, 1, 11, 63

    If mont_date_received = "__ __ __" Then
        mont_date_received = "NONE"
    Else
        mont_date_received = replace(mont_date_received, " ", "/")
    End If
    If mont_cash_status = "I" Then mont_cash_status = "I - HRF Received - INCOMPLETE"
    If mont_cash_status = "U" Then mont_cash_status = "U - HRF Received - COMPLETE"
    If mont_cash_status = "A" Then mont_cash_status = "A - HRF Processed - APPROVED"
    If mont_cash_status = "O" Then mont_cash_status = "O - Overrid Autoclose"
    If mont_cash_status = "T" Then mont_cash_status = "T - No HRF - TERMINATED"
    If mont_cash_status = "N" Then mont_cash_status = "N - HRF Not Received"
    If mont_cash_status = "D" Then mont_cash_status = "D - FS Denied"
    If mont_cash_status = "_" Then mont_cash_status = "NONE"

    If mont_snap_status = "I" Then mont_snap_status = "I - HRF Received - INCOMPLETE"
    If mont_snap_status = "U" Then mont_snap_status = "U - HRF Received - COMPLETE"
    If mont_snap_status = "A" Then mont_snap_status = "A - HRF Processed - APPROVED"
    If mont_snap_status = "O" Then mont_snap_status = "O - Overrid Autoclose"
    If mont_snap_status = "T" Then mont_snap_status = "T - No HRF - TERMINATED"
    If mont_snap_status = "N" Then mont_snap_status = "N - HRF Not Received"
    If mont_snap_status = "D" Then mont_snap_status = "D - FS Denied"
    If mont_snap_status = "_" Then mont_snap_status = "NONE"

    If mont_hc_status = "I" Then mont_hc_status = "I - HRF Received - INCOMPLETE"
    If mont_hc_status = "U" Then mont_hc_status = "U - HRF Received - COMPLETE"
    If mont_hc_status = "A" Then mont_hc_status = "A - HRF Processed - APPROVED"
    If mont_hc_status = "O" Then mont_hc_status = "O - Overrid Autoclose"
    If mont_hc_status = "T" Then mont_hc_status = "T - No HRF - TERMINATED"
    If mont_hc_status = "N" Then mont_hc_status = "N - HRF Not Received"
    If mont_hc_status = "D" Then mont_hc_status = "D - FS Denied"
    If mont_hc_status = "_" Then mont_hc_status = "NONE"

    mont_detail = "This case has monthly reporting.    Date HRF Received: " & mont_date_received & "      Cash Status: " & mont_cash_status & "     SNAP Status: " & mont_snap_status & "     HC Status: " & mont_hc_status
End If

income_counter = 0
asset_counter = 0
For each member in HH_member_array
    Call navigate_to_MAXIS_screen("STAT", "JOBS")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_INCOME_ARRAY(update_checkbox, income_counter)

            ALL_INCOME_ARRAY(category_const, income_counter) = "JOBS"
            Call access_JOBS_panel("READ", ALL_INCOME_ARRAY(owner_name, income_counter), ALL_INCOME_ARRAY(verif_const, income_counter), ALL_INCOME_ARRAY(name_const, income_counter), ALL_INCOME_ARRAY(type_const, income_counter), ALL_INCOME_ARRAY(pay_amt_const, income_counter), ALL_INCOME_ARRAY(amount_const, income_counter), ALL_INCOME_ARRAY(hours_const, income_counter), ALL_INCOME_ARRAY(frequency_const, income_counter),  ALL_INCOME_ARRAY(update_date_const, income_counter), ALL_INCOME_ARRAY(start_date_const, income_counter), ALL_INCOME_ARRAY(end_date_const, income_counter), ALL_INCOME_ARRAY(panel_instance, income_counter), ALL_INCOME_ARRAY(jobs_hrly_wage, income_counter), ALL_INCOME_ARRAY(retro_income_amount, income_counter), ALL_INCOME_ARRAY(retro_income_hours, income_counter), ALL_INCOME_ARRAY(snap_pic_frequency, income_counter), ALL_INCOME_ARRAY(snap_pic_hours_per_pay, income_counter), ALL_INCOME_ARRAY(snap_pic_income_per_pay, income_counter), ALL_INCOME_ARRAY(snap_pic_monthly_income, income_counter), ALL_INCOME_ARRAY(grh_pic_frequency, income_counter), ALL_INCOME_ARRAY(grh_pic_income_per_pay, income_counter), ALL_INCOME_ARRAY(grh_pic_monthly_income, income_counter), ALL_INCOME_ARRAY(jobs_subsidy, income_counter))
            ALL_INCOME_ARRAY(owner_ref_const, income_counter) = ALL_INCOME_ARRAY(owner_name, income_counter)
            For each person in full_hh_list
                ' MsgBox "Person - " & person & vbNewLine & "LEFT, 2 - " & left(person, 2) & vbNewLine & "Income Owner - " & ALL_INCOME_ARRAY(owner_name, income_counter)
                If left(person, 2) = ALL_INCOME_ARRAY(owner_name, income_counter) Then ALL_INCOME_ARRAY(owner_name, income_counter) = person
            Next
            ' MsgBox "Line 974" & vbNewLine & ALL_INCOME_ARRAY(update_date_const, income_counter)
            ALL_INCOME_ARRAY(update_date_const, income_counter) = DateValue(ALL_INCOME_ARRAY(update_date_const, income_counter))
            If DateDiff("d", ALL_INCOME_ARRAY(update_date_const, income_counter), date) = 0 Then ALL_INCOME_ARRAY(new_checkbox, income_counter) = checked

            income_counter = income_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "BUSI")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    ' MsgBox versions
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_INCOME_ARRAY(update_checkbox, income_counter)

            ALL_INCOME_ARRAY(category_const, income_counter) = "BUSI"
            Call access_BUSI_panel("READ", ALL_INCOME_ARRAY(owner_name, income_counter), ALL_INCOME_ARRAY(type_const, income_counter), ALL_INCOME_ARRAY(start_date_const, income_counter), ALL_INCOME_ARRAY(end_date_const, income_counter), ALL_INCOME_ARRAY(busi_cash_net_prosp, income_counter), ALL_INCOME_ARRAY(busi_cash_net_retro, income_counter), ALL_INCOME_ARRAY(busi_cash_gross_retro, income_counter), ALL_INCOME_ARRAY(busi_cash_expense_retro, income_counter), ALL_INCOME_ARRAY(busi_cash_gross_prosp, income_counter), ALL_INCOME_ARRAY(busi_cash_expense_prosp, income_counter), ALL_INCOME_ARRAY(busi_cash_income_verif, income_counter), ALL_INCOME_ARRAY(busi_cash_expense_verif, income_counter), ALL_INCOME_ARRAY(busi_snap_net_prosp, income_counter), ALL_INCOME_ARRAY(busi_snap_net_retro, income_counter), ALL_INCOME_ARRAY(busi_snap_gross_retro, income_counter), ALL_INCOME_ARRAY(busi_snap_expense_retro, income_counter), ALL_INCOME_ARRAY(busi_snap_gross_prosp, income_counter), ALL_INCOME_ARRAY(busi_snap_expense_prosp, income_counter), ALL_INCOME_ARRAY(busi_snap_income_verif, income_counter), ALL_INCOME_ARRAY(busi_snap_expense_verif, income_counter), ALL_INCOME_ARRAY(busi_hc_a_net_prosp, income_counter), ALL_INCOME_ARRAY(busi_hc_b_net_prosp, income_counter), ALL_INCOME_ARRAY(busi_hc_a_gross_prosp, income_counter), ALL_INCOME_ARRAY(busi_hc_a_expense_prosp, income_counter), ALL_INCOME_ARRAY(busi_hc_a_income_verif, income_counter), ALL_INCOME_ARRAY(busi_hc_a_expense_verif, income_counter), ALL_INCOME_ARRAY(busi_hc_b_gross_prosp, income_counter), ALL_INCOME_ARRAY(busi_hc_b_expense_prosp, income_counter), ALL_INCOME_ARRAY(busi_hc_b_income_verif, income_counter), ALL_INCOME_ARRAY(busi_hc_b_expense_verif, income_counter), ALL_INCOME_ARRAY(busi_se_method, income_counter), ALL_INCOME_ARRAY(busi_se_method_date, income_counter), ALL_INCOME_ARRAY(rptd_hours_const, income_counter), ALL_INCOME_ARRAY(min_wg_hours_const, income_counter), ALL_INCOME_ARRAY(update_date_const, income_counter), ALL_INCOME_ARRAY(panel_instance, income_counter))
            ALL_INCOME_ARRAY(owner_ref_const, income_counter) = ALL_INCOME_ARRAY(owner_name, income_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_INCOME_ARRAY(owner_name, income_counter) Then ALL_INCOME_ARRAY(owner_name, income_counter) = person
            Next
            ' MsgBox "Line 1000" & vbNewLine & ALL_INCOME_ARRAY(update_date_const, income_counter)
            ALL_INCOME_ARRAY(update_date_const, income_counter) = DateValue(ALL_INCOME_ARRAY(update_date_const, income_counter))

            income_counter = income_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "UNEA")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_INCOME_ARRAY(update_checkbox, income_counter)


            ALL_INCOME_ARRAY(category_const, income_counter) = "UNEA"
            Call access_UNEA_panel("READ", ALL_INCOME_ARRAY(owner_name, income_counter), ALL_INCOME_ARRAY(type_const, income_counter), ALL_INCOME_ARRAY(verif_const, income_counter), ALL_INCOME_ARRAY(claim_nbr_const, income_counter), ALL_INCOME_ARRAY(start_date_const, income_counter), ALL_INCOME_ARRAY(end_date_const, income_counter), ALL_INCOME_ARRAY(cola_disregard_amt, income_counter), ALL_INCOME_ARRAY(amount_const, income_counter), ALL_INCOME_ARRAY(pay_amt_const, income_counter), ALL_INCOME_ARRAY(frequency_const, income_counter), ALL_INCOME_ARRAY(update_date_const, income_counter), ALL_INCOME_ARRAY(panel_instance, income_counter),             ALL_INCOME_ARRAY(snap_pic_income_per_pay, income_counter), ALL_INCOME_ARRAY(snap_pic_monthly_income, income_counter), ALL_INCOME_ARRAY(retro_income_amount, income_counter))
            ALL_INCOME_ARRAY(owner_ref_const, income_counter) = ALL_INCOME_ARRAY(owner_name, income_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_INCOME_ARRAY(owner_name, income_counter) Then ALL_INCOME_ARRAY(owner_name, income_counter) = person
            Next
            ' MsgBox "Line 1000" & vbNewLine & ALL_INCOME_ARRAY(update_date_const, income_counter)
            ALL_INCOME_ARRAY(update_date_const, income_counter) = DateValue(ALL_INCOME_ARRAY(update_date_const, income_counter))

            income_counter = income_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "CASH")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_ASSETS_ARRAY(update_checkbox, asset_counter)

            ALL_ASSETS_ARRAY(category_const, asset_counter) = "CASH"
            EMReadScreen ALL_ASSETS_ARRAY(owner_name, asset_counter), 2, 4, 33
            ALL_ASSETS_ARRAY(owner_ref_const, asset_counter) = ALL_ASSETS_ARRAY(owner_name, asset_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_ASSETS_ARRAY(owner_name, asset_counter) Then ALL_ASSETS_ARRAY(owner_name, asset_counter) = person
            Next
            ALL_ASSETS_ARRAY(type_const, asset_counter) = "CASH"
            EMReadScreen ALL_ASSETS_ARRAY(amount_const, asset_counter), 8, 8, 39
            ALL_ASSETS_ARRAY(amount_const, asset_counter) = trim(ALL_ASSETS_ARRAY(amount_const, asset_counter))

            asset_counter = asset_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "ACCT")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_ASSETS_ARRAY(update_checkbox, asset_counter)

            ALL_ASSETS_ARRAY(category_const, asset_counter) = "ACCT"

            Call access_ACCT_panel("READ", ALL_ASSETS_ARRAY(owner_name, asset_counter), ALL_ASSETS_ARRAY(type_const, asset_counter), ALL_ASSETS_ARRAY(id_number_const,asset_counter), ALL_ASSETS_ARRAY(name_const, asset_counter), ALL_ASSETS_ARRAY(amount_const, asset_counter), ALL_ASSETS_ARRAY(verif_const, asset_counter), ALL_ASSETS_ARRAY(update_date_const, asset_counter), ALL_ASSETS_ARRAY(panel_instance, asset_counter), ALL_ASSETS_ARRAY(balance_date_const, asset_counter), ALL_ASSETS_ARRAY(withdraw_penalty_const, asset_counter), ALL_ASSETS_ARRAY(withdraw_yn_const, asset_counter), ALL_ASSETS_ARRAY(withdraw_verif_const, asset_counter), ALL_ASSETS_ARRAY(count_cash_const, asset_counter), ALL_ASSETS_ARRAY(count_snap_const, asset_counter), ALL_ASSETS_ARRAY(count_hc_const, asset_counter), ALL_ASSETS_ARRAY(count_grh_const, asset_counter), ALL_ASSETS_ARRAY(count_ive_const, asset_counter), ALL_ASSETS_ARRAY(joint_own_const, asset_counter), ALL_ASSETS_ARRAY(share_ratio_const, asset_counter), ALL_ASSETS_ARRAY(next_interst_const, asset_counter))

            ALL_ASSETS_ARRAY(owner_ref_const, asset_counter) = ALL_ASSETS_ARRAY(owner_name, asset_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_ASSETS_ARRAY(owner_name, asset_counter) Then ALL_ASSETS_ARRAY(owner_name, asset_counter) = person
            Next

            asset_counter = asset_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "SECU")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_ASSETS_ARRAY(update_checkbox, asset_counter)

            ALL_ASSETS_ARRAY(category_const, asset_counter) = "SECU"
            Call access_SECU_panel("READ", ALL_ASSETS_ARRAY(owner_name, asset_counter), ALL_ASSETS_ARRAY(type_const, asset_counter), ALL_ASSETS_ARRAY(id_number_const, asset_counter), ALL_ASSETS_ARRAY(name_const, asset_counter), ALL_ASSETS_ARRAY(amount_const, asset_counter), ALL_ASSETS_ARRAY(verif_const, asset_counter), ALL_ASSETS_ARRAY(update_date_const, asset_counter), ALL_ASSETS_ARRAY(panel_instance, asset_counter), ALL_ASSETS_ARRAY(face_value_const, asset_counter), ALL_ASSETS_ARRAY(withdraw_penalty_const, asset_counter), ALL_ASSETS_ARRAY(withdraw_yn_const, asset_counter), ALL_ASSETS_ARRAY(withdraw_verif_const, asset_counter), ALL_ASSETS_ARRAY(count_cash_const, asset_counter), ALL_ASSETS_ARRAY(count_snap_const, asset_counter), ALL_ASSETS_ARRAY(count_hc_const, asset_counter), ALL_ASSETS_ARRAY(count_grh_const, asset_counter), ALL_ASSETS_ARRAY(count_ive_const, asset_counter), ALL_ASSETS_ARRAY(joint_own_const, asset_counter), ALL_ASSETS_ARRAY(share_ratio_const, asset_counter), ALL_ASSETS_ARRAY(balance_date_const, asset_counter))
            ALL_ASSETS_ARRAY(owner_ref_const, asset_counter) = ALL_ASSETS_ARRAY(owner_name, asset_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_ASSETS_ARRAY(owner_name, asset_counter) Then ALL_ASSETS_ARRAY(owner_name, asset_counter) = person
            Next

            asset_counter = asset_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "CARS")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_ASSETS_ARRAY(update_checkbox, asset_counter)



            ALL_ASSETS_ARRAY(category_const, asset_counter) = "CARS"
            Call access_CARS_panel("READ", ALL_ASSETS_ARRAY(owner_name, asset_counter), ALL_ASSETS_ARRAY(type_const, asset_counter), ALL_ASSETS_ARRAY(year_const, asset_counter), ALL_ASSETS_ARRAY(make_const, asset_counter), ALL_ASSETS_ARRAY(model_const, asset_counter), ALL_ASSETS_ARRAY(verif_const, asset_counter), ALL_ASSETS_ARRAY(update_date_const, asset_counter), ALL_ASSETS_ARRAY(panel_instance, asset_counter), ALL_ASSETS_ARRAY(trade_in_const, asset_counter), ALL_ASSETS_ARRAY(loan_const, asset_counter), ALL_ASSETS_ARRAY(source_const, asset_counter), ALL_ASSETS_ARRAY(owed_amt_const, asset_counter), ALL_ASSETS_ARRAY(owed_verif_const, asset_counter), ALL_ASSETS_ARRAY(owed_date_const, asset_counter), ALL_ASSETS_ARRAY(cars_use_const, asset_counter), ALL_ASSETS_ARRAY(hc_benefit_const, asset_counter), ALL_ASSETS_ARRAY(joint_own_const, asset_counter), ALL_ASSETS_ARRAY(share_ratio_const, asset_counter))
            ALL_ASSETS_ARRAY(owner_ref_const, asset_counter) = ALL_ASSETS_ARRAY(owner_name, asset_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_ASSETS_ARRAY(owner_name, asset_counter) Then ALL_ASSETS_ARRAY(owner_name, asset_counter) = person
            Next
            ALL_ASSETS_ARRAY(make_model_yr, asset_counter) = ALL_ASSETS_ARRAY(year_const, asset_counter) & " " & ALL_ASSETS_ARRAY(make_const, asset_counter) & " " & ALL_ASSETS_ARRAY(model_const, asset_counter)

            asset_counter = asset_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If

    Call navigate_to_MAXIS_screen("STAT", "REST")
    EMWriteScreen member, 20, 76
    transmit
    EMReadScreen versions, 6, 2, 73
    If versions <> "0 Of 0" Then
        Do
            ReDim Preserve ALL_ASSETS_ARRAY(update_checkbox, asset_counter)

            ALL_ASSETS_ARRAY(category_const, asset_counter) = "REST"
            Call access_REST_panel("READ", ALL_ASSETS_ARRAY(owner_name, asset_counter), ALL_ASSETS_ARRAY(type_const, asset_counter), ALL_ASSETS_ARRAY(verif_const, asset_counter), ALL_ASSETS_ARRAY(update_date_const, asset_counter), ALL_ASSETS_ARRAY(panel_instance, asset_counter),             ALL_ASSETS_ARRAY(market_value_const, asset_counter), ALL_ASSETS_ARRAY(value_verif_const, asset_counter), ALL_ASSETS_ARRAY(owed_amt_const, asset_counter), ALL_ASSETS_ARRAY(owed_verif_const, asset_counter), ALL_ASSETS_ARRAY(owed_date_const, asset_counter), ALL_ASSETS_ARRAY(rest_prop_status_const, asset_counter), ALL_ASSETS_ARRAY(joint_own_const, asset_counter), ALL_ASSETS_ARRAY(share_ratio_const, asset_counter), ALL_ASSETS_ARRAY(rest_repymt_date_const, asset_counter))
            ALL_ASSETS_ARRAY(owner_ref_const, asset_counter) = ALL_ASSETS_ARRAY(owner_name, asset_counter)
            For each person in full_hh_list
                If left(person, 2) = ALL_ASSETS_ARRAY(owner_name, asset_counter) Then ALL_ASSETS_ARRAY(owner_name, asset_counter) = person
            Next

            asset_counter = asset_counter + 1
            transmit
            EMReadScreen last_panel, 7, 24, 2
        Loop until last_panel = "ENTER A"
    End If
Next

For case_panels = 0 to UBound(ALL_PANELS_ARRAY, 2)
    If ALL_PANELS_ARRAY(one_per_case_const, case_panels) = TRUE Then
    '"ADDR" "AREP" "ALTP" "TYPE" "PROG" "HCRE" "EATS" "SIBL" "DSTT" "HEST" "PACT" "SWKR" "REVW" "MISC" "RESI" "BILS" "BUDG" "MMSA"

    ElseIf ALL_PANELS_ARRAY(one_per_person_const, case_panels) = TRUE Then
    '"MEMB" "MEMI" "ALIA" "PARE" "IMIG" "SPON" "ADME" "REMO" "DISA" "PREG" "STRK" "STWK" "SCHL" "WREG" "EMPS" "STIN" "STEC" "CASH"
    '"PBEN" "LUMP" "TRAC" "DCEX" "WKEX" "COEX" "SHEL" "ACUT" "PDED" "FMED" "MEDI" "DIET" "TIME" "EMMA" "HCMI" "SANC" "DFLN" "MSUR" "SSRT"
        If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "MEMB" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "MEMI" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "WREG" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "SHEL" Then
            For pers_info = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
                ' MsgBox "all clts - " & ALL_CLIENTS_ARRAY(memb_ref_numb, pers_info) & vbNewLine & "all panels - " & ALL_PANELS_ARRAY(the_memb_const, case_panels)
                If ALL_CLIENTS_ARRAY(memb_ref_numb, pers_info) = ALL_PANELS_ARRAY(the_memb_const, case_panels) Then
                    ALL_PANELS_ARRAY(array_ref_const, case_panels) = pers_info
                    ' MsgBox "MATCH!" & vbNewLine & "The ref saved - " & ALL_PANELS_ARRAY(array_ref_const, case_panels)
                End If
            Next
        ElseIF ALL_PANELS_ARRAY(the_panel_const, case_panels) = "CASH" Then
            For acct_panel = 0 to UBound(ALL_ASSETS_ARRAY, 2)
                ' MsgBox "Asset category - " & ALL_ASSETS_ARRAY(category_const, acct_panel) & " --- Panel category - " & ALL_PANELS_ARRAY(the_panel_const, case_panels) & vbNewLine & "Asset owner - " & ALL_ASSETS_ARRAY(owner_ref_const, acct_panel) & " --- Panel owner - " & ALL_PANELS_ARRAY(the_memb_const, case_panels) & vbNewLine & "Asset instance - " & ALL_ASSETS_ARRAY(panel_instance, acct_panel) & " --- Panel instance - " &  ALL_PANELS_ARRAY(the_inst_const, case_panels)

                If ALL_ASSETS_ARRAY(category_const, acct_panel) = ALL_PANELS_ARRAY(the_panel_const, case_panels) AND ALL_ASSETS_ARRAY(owner_ref_const, acct_panel) = ALL_PANELS_ARRAY(the_memb_const, case_panels) AND ALL_PANELS_ARRAY(the_inst_const, case_panels) = ALL_ASSETS_ARRAY(panel_instance, acct_panel) Then
                    ALL_PANELS_ARRAY(array_ref_const, case_panels) = acct_panel
                    ' MsgBox "MATCH!" & vbNewLine & "The ref saved - " & ALL_PANELS_ARRAY(array_ref_const, case_panels)
                End If
            Next
        End If
    ElseIf ALL_PANELS_ARRAY(multiple_per_case, case_panels) = TRUE Then
    '"FCFC" "FCPL" "ABPS" "INSA"

    ElseIf ALL_PANELS_ARRAY(multiple_per_person_const, case_panels) = TRUE Then
    '"FACI" "ACCT" "SECU" "CARS" "REST" "OTHR" "TRAN" "UNEA" "RBIC" "BUSI" "JOBS" "ACCI" "DISQ"
        If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "JOBS" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "BUSI" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "UNEA" Then
            For ei_panel = 0 to UBound(ALL_INCOME_ARRAY, 2)
                ' MsgBox "all jobs - " & ALL_INCOME_ARRAY(owner_ref_const, ei_panel) & vbNewLine & "all panels - " & ALL_PANELS_ARRAY(the_memb_const, case_panels)
                If ALL_INCOME_ARRAY(category_const, ei_panel) = ALL_PANELS_ARRAY(the_panel_const, case_panels) AND ALL_INCOME_ARRAY(owner_ref_const, ei_panel) = ALL_PANELS_ARRAY(the_memb_const, case_panels) AND ALL_PANELS_ARRAY(the_inst_const, case_panels) = ALL_INCOME_ARRAY(panel_instance, ei_panel) Then
                    ALL_PANELS_ARRAY(array_ref_const, case_panels) = ei_panel
                    ' MsgBox "MATCH!" & vbNewLine & "The ref saved - " & ALL_PANELS_ARRAY(array_ref_const, case_panels)
                End If
            Next
        ElseIF ALL_PANELS_ARRAY(the_panel_const, case_panels) = "CASH" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "ACCT" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "SECU" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "REST" OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = "CARS" Then
            For acct_panel = 0 to UBound(ALL_ASSETS_ARRAY, 2)
                ' MsgBox "Asset category - " & ALL_ASSETS_ARRAY(category_const, acct_panel) & " --- Panel category - " & ALL_PANELS_ARRAY(the_panel_const, case_panels) & vbNewLine & "Asset owner - " & ALL_ASSETS_ARRAY(owner_ref_const, acct_panel) & " --- Panel owner - " & ALL_PANELS_ARRAY(the_memb_const, case_panels) & vbNewLine & "Asset instance - " & ALL_ASSETS_ARRAY(panel_instance, acct_panel) & " --- Panel instance - " &  ALL_PANELS_ARRAY(the_inst_const, case_panels)

                If ALL_ASSETS_ARRAY(category_const, acct_panel) = ALL_PANELS_ARRAY(the_panel_const, case_panels) AND ALL_ASSETS_ARRAY(owner_ref_const, acct_panel) = ALL_PANELS_ARRAY(the_memb_const, case_panels) AND ALL_PANELS_ARRAY(the_inst_const, case_panels) = ALL_ASSETS_ARRAY(panel_instance, acct_panel) Then
                    ALL_PANELS_ARRAY(array_ref_const, case_panels) = acct_panel
                    ' MsgBox "MATCH!" & vbNewLine & "The ref saved - " & ALL_PANELS_ARRAY(array_ref_const, case_panels)
                End If
            Next
        ElseIF ALL_PANELS_ARRAY(the_panel_const, case_panels) = "FACI" Then
            For faci_panel = 0 to UBound(FACILITIES_ARRAY, 2)
                If FACILITIES_ARRAY(faci_ref_numb, faci_panel) = ALL_PANELS_ARRAY(the_memb_const, case_panels) AND ALL_PANELS_ARRAY(the_inst_const, case_panels) = FACILITIES_ARRAY(faci_instance, faci_panel) Then
                    ALL_PANELS_ARRAY(array_ref_const, case_panels) = faci_panel
                End If
            Next
        End If
    End If
Next

'NEED TO Address

'Does the residence address match?
'Does the mailing address match?
'If phone number is listed in dialog - check to see if the same number is on addr.

'If someone has moved in or out
'Look through each member to see if marked as moved in or out.
'if a person is listed as not added in MAXIS yet - take notes of this person information

'Q4 - MA - apply for someone else
    'Has it been left blank?
    'If others are requesting - note information for each member.

'Q5 - MA - self-employed
    'Has it been left blank?
    'Look through each self-employment and if marked as new - add dialog to collect information
    'Look through each self-employment and if detail is marked, add dialog to collect information

'Q6 - MA - working
    'Has it been left blank?
    'Look through each job and if marked as new - add dialog to collect information
    'Look through each job and if detail is marked, add dialog to collect information

'Q7 - MA - unearned income
    'Has it been left blank?
    'Look through each unea and if marked as new - add dialog to collect information
    'Look through each unea and if detail is marked, add dialog to collect information

'Q9 - MA - liquid assets
    'Has it been left blank?
    'Look through each liquid asset and if marked as new - add dialog to collect information
    'Look through each liquid asset and if detail is marked, add dialog to collect information

'Q10 - MA - Securities
    'Has it been left blank?
    'Look through each security and if marked as new - add dialog to collect information
    'Look through each security and if detail is marked, add dialog to collect information

'Q11 - MA - vehicle
    'Has it been left blank?
    'Look through each vehicle and if marked as new - add dialog to collect information
    'Look through each vehicle and if detail is marked, add dialog to collect information

'Q12 - MA - real estate
    'Has it been left blank?
    'Look through each real estate and if marked as new - add dialog to collect information
    'Look through each real estate and if detail is marked, add dialog to collect information

'Q13 - MA - other changes
    'no additional dialog needed here '

'Q15 - SNAP - household moved
    'If yes - dialog with more detail
    '

'Q16 - SNAP - Change in Earned Income
    'If a job is listed, see if yuo can find it in MAXIS
    'Dialog the information/detail

'Q17 - SNAP - Change in Unearned Income
    'If earned income is listed - see if we can find it in MAXIS
    'Dialog the informaiton/detail

'Q18 - SNAP - Change in Child Support
    'if child support is indicated, try to find the panel
    'dialog information

'Q19 - SNAP - ABAWD Hours
'   'Review the abawd information'

verifs_needed_for_SNAP_checkbox = checked
verifs_needed_for_HC_checkbox = checked
verifs_needed_for_GRH_checkbox = checked

For new_income = 0 to UBound(NEW_EARNED_ARRAY, 2)
    If NEW_EARNED_ARRAY(earned_client, new_income) = "Select or Type" AND trim(NEW_EARNED_ARRAY(earned_source, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_start_date, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_amount, new_income)) = "" Then
        NEW_EARNED_ARRAY(earned_type, new_income) = ""
        NEW_EARNED_ARRAY(earned_prog_list, new_income) = ""
    End If
    If NEW_EARNED_ARRAY(earned_client, new_income) = "Select or Type" AND trim(NEW_EARNED_ARRAY(earned_source, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_start_date, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_seasonal, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_amount, new_income)) = "" AND NEW_EARNED_ARRAY(earned_freq, new_income) = "Select One..." Then
        NEW_EARNED_ARRAY(earned_type, new_income) = ""
        NEW_EARNED_ARRAY(earned_prog_list, new_income) = ""
    End If
    If NEW_EARNED_ARRAY(earned_client, new_income) = "Select or Type" AND trim(NEW_EARNED_ARRAY(earned_source, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_change_date, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_hours, new_income)) = "" AND trim(NEW_EARNED_ARRAY(earned_amount, new_income)) = "" AND NEW_EARNED_ARRAY(earned_freq, new_income) = "Select One..." Then
        NEW_EARNED_ARRAY(earned_type, new_income) = ""
        NEW_EARNED_ARRAY(earned_prog_list, new_income) = ""
    End If
Next
For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
    If NEW_UNEARNED_ARRAY(unearned_client, each_unea) = "Select or Type" AND NEW_UNEARNED_ARRAY(unearned_source, each_unea) = "Select or Type" AND trim(NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)) = "" AND trim(NEW_UNEARNED_ARRAY(unearned_amount, each_unea)) = "" AND NEW_UNEARNED_ARRAY(unearned_freq, each_unea) = "Select One..." Then
        NEW_UNEARNED_ARRAY(unearned_type, each_unea) = ""
        NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = ""
    End If
    If NEW_UNEARNED_ARRAY(unearned_client, each_unea) = "Select or Type" AND NEW_UNEARNED_ARRAY(unearned_source, each_unea) = "Select or Type" AND trim(NEW_UNEARNED_ARRAY(unearned_change_date, each_unea)) = "" AND trim(NEW_UNEARNED_ARRAY(unearned_amount, each_unea)) = "" AND NEW_UNEARNED_ARRAY(unearned_freq, each_unea) = "Select One..." Then
        NEW_UNEARNED_ARRAY(unearned_type, each_unea) = ""
        NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = ""
    End If
Next
For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
    If NEW_ASSET_ARRAY(asset_client, each_asset) = "Select or Type" AND NEW_ASSET_ARRAY(asset_acct_type, each_asset) = "Select or Type" AND trim(NEW_ASSET_ARRAY(asset_bank_name, each_asset)) = "" Then
        NEW_ASSET_ARRAY(asset_type, each_asset) = ""
        NEW_ASSET_ARRAY(asset_prog_list, each_asset) = ""
    End If
    If NEW_ASSET_ARRAY(asset_client, each_asset) = "Select or Type" AND NEW_ASSET_ARRAY(asset_acct_type, each_asset) = "Select or Type" AND trim(NEW_ASSET_ARRAY(asset_year_make_model, each_asset)) = "" Then
        NEW_ASSET_ARRAY(asset_type, each_asset) = ""
        NEW_ASSET_ARRAY(asset_prog_list, each_asset) = ""
    End If
    If NEW_ASSET_ARRAY(asset_client, each_asset) = "Select or Type" AND NEW_ASSET_ARRAY(asset_acct_type, each_asset) = "Select or Type" AND trim(NEW_ASSET_ARRAY(asset_address, each_asset)) = "" Then
        NEW_ASSET_ARRAY(asset_type, each_asset) = ""
        NEW_ASSET_ARRAY(asset_prog_list, each_asset) = ""
    End If
Next
'IF SNAP SR - if complete then the details, if incomplete - then why (Signed too early, not all people required signed, no proofs, not all questions answered)'
'IF GRH SR - if complete then the details, if incomplete - then why (Signed too early, not all people required signed, no proofs, not all questions answered)'
'IF HC SR - if complete then the details, if incomplete - then why (Signed too early, not all people required signed, no proofs, not all questions answered)'
addr_panel_count = 0
memb_panel_count = 0
wreg_panel_count = 0
faci_panel_count = 0
revw_panel_count = 0
acct_panel_count = 0
acct_panel_count = 0
secu_panel_count = 0
cars_panel_count = 0
rest_panel_count = 0
busi_panel_count = 0
jobs_panel_count = 0
unea_panel_count = 0
shel_panel_count = 0
hest_panel_count = 0
For case_panels = 0 to UBound(ALL_PANELS_ARRAY, 2)
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "ADDR" Then addr_panel_count = addr_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "MEMB" Then memb_panel_count = memb_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "WREG" Then wreg_panel_count = wreg_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "FACI" Then faci_panel_count = faci_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "REVW" Then revw_panel_count = revw_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "ACCT" Then acct_panel_count = acct_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "CASH" Then acct_panel_count = acct_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "SECU" Then secu_panel_count = secu_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "CARS" Then cars_panel_count = cars_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "REST" Then rest_panel_count = rest_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "BUSI" Then busi_panel_count = busi_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "JOBS" Then jobs_panel_count = jobs_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "UNEA" Then unea_panel_count = unea_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "SHEL" Then shel_panel_count = shel_panel_count + 1
    If ALL_PANELS_ARRAY(the_panel_const, case_panels) = "HEST" Then hest_panel_count = hest_panel_count + 1
Next
address_button_name = "ADDRESS (" & addr_panel_count & ")"
memb_button_name = "HH COMP (" & memb_panel_count & ")"
wreg_button_name = "WREG (" & wreg_panel_count & ")"
faci_button_name = "FACI (" & faci_panel_count & ")"
revw_button_name = "REVW (" & revw_panel_count & ")"
assets_button_name = "LIQUID ASSETS (" & acct_panel_count & ")"
secu_button_name = "SECURITY (" & secu_panel_count & ")"
cars_button_name = "VEHICLES (" & cars_panel_count & ")"
rest_button_name = "REAL ESTATE (" & rest_panel_count & ")"
busi_nutton_name = "BUSI (" & busi_panel_count & ")"
jobs_button_name = "JOBS (" & jobs_panel_count & ")"
unea_button_name = "UNEA (" & unea_panel_count & ")"
shel_button_name = "SHELTER (" & shel_panel_count & ")"
hest_button_name = "UTILITIES (" & hest_panel_count & ")"

Call generate_client_list(verification_memb_list, "Select or Type Member")
verification_memb_list = " "+chr(9)+verification_memb_list
If residence_address_match_yn = "No - there is a difference." OR mailing_address_match_yn = "No - there is a difference." Then update_addr_checkbox = checked

new_resi_state = "MN Minnesota"
new_mail_state = "MN Minnesota"
new_resi_county = "27 Hennepin"

notes_reviewed = FALSE
verif_reviewed = FALSE

form_note = 1
stat_note = 2
verif_note = 3
note_to_show = form_note

panel_indicator = "ADDR"
panel_array_to_use = 0
show_all_the_panels = TRUE
Do
    Do
        If show_all_the_panels = TRUE Then
            Do
                Do
                    err_msg = ""

                    If panel_indicator = "ACCT" Then second_indicator = "CASH"
                    If panel_indicator = "CASH" Then
                        panel_indicator = "ACCT"
                        second_indicator = "CASH"
                    End If
                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 605, 370, "CSR Information Details"
                      ButtonGroup ButtonPressed
                        PushButton 430, 355, 50, 10, "NEXT", next_btn
                        If panel_indicator = "ADDR" Then
                            Text 525, 8, 75, 15, "         " & address_button_name
                        Else
                            PushButton 525, 5, 75, 15, address_button_name, address_tab_btn
                        End If
                        If panel_indicator = "MEMB" Then
                            Text 525, 22, 75, 15, "        " & memb_button_name
                        Else
                            PushButton 525, 20, 75, 15, memb_button_name, hh_comp_tab_btn
                        End If
                        If panel_indicator = "WREG" Then
                            Text 525, 37, 75, 15, "          " & wreg_button_name
                        Else
                            PushButton 525, 35, 75, 15, wreg_button_name, wreg_btn
                        End If
                        If panel_indicator = "FACI" Then
                            Text 525, 52, 75, 15, "            " & faci_button_name
                        Else
                            PushButton 525, 50, 75, 15, faci_button_name, faci_btn
                        End If
                        If panel_indicator = "REVW" Then
                            Text 525, 67, 75, 15, "          " & revw_button_name
                        Else
                            PushButton 525, 65, 75, 15, revw_button_name, revw_btn
                        End If
                        If panel_indicator = "ACCT" Then
                            Text 525, 82, 75, 15, "  " & assets_button_name
                        Else
                            PushButton 525, 80, 75, 15, assets_button_name, acct_tab_btn
                        End If
                        If panel_indicator = "SECU" Then
                            Text 525, 97, 75, 15, "        " & secu_button_name
                        Else
                            PushButton 525, 95, 75, 15, secu_button_name, secu_tab_btn
                        End If
                        If panel_indicator = "CARS" Then
                            Text 525, 112, 75, 15, "        " & cars_button_name
                        Else
                            PushButton 525, 110, 75, 15, cars_button_name, cars_tab_btn
                        End If
                        If panel_indicator = "REST" Then
                            Text 525, 127, 75, 15, "     " & rest_button_name
                        Else
                            PushButton 525, 125, 75, 15, rest_button_name, rest_tab_btn
                        End If
                        If panel_indicator = "UNEA" Then
                            Text 525, 142, 75, 15, "           " & unea_button_name
                        Else
                            PushButton 525, 140, 75, 15, unea_button_name, unea_tab_btn
                        End If
                        If panel_indicator = "BUSI" Then
                            Text 525, 157, 75, 15, "            " & busi_nutton_name
                        Else
                            PushButton 525, 155, 75, 15, busi_nutton_name, busi_tabs_btn
                        End If
                        If panel_indicator = "JOBS" Then
                            Text 525, 172, 75, 15, "           " & jobs_button_name
                        Else
                            PushButton 525, 170, 75, 15, jobs_button_name, jobs_tab_btn
                        End If
                        If panel_indicator = "SHEL" Then
                            Text 525, 187, 75, 15, "        " & shel_button_name
                        Else
                            PushButton 525, 185, 75, 15, shel_button_name, shel_tab_btn
                        End If
                        If panel_indicator = "HEST" Then
                            Text 525, 202, 75, 15, "         " & hest_button_name
                        Else
                            PushButton 525, 200, 75, 15, hest_button_name, hst_tab_btn
                        End If
                        If panel_indicator = "VERIFS" Then
                            Text 525, 232, 75, 15, "     VERIFICATIONS"
                        Else
                            PushButton 525, 230, 75, 15, "VERIFICATIONS", verifs_tab_btn
                        End If
                        If panel_indicator = "NOTES" Then
                            Text 525, 247, 75, 15, "              NOTES"
                        Else
                            PushButton 525, 245, 75, 15, "NOTES", notes_tab_btn
                        End If

                        x_pos = 10
                        For case_panels = 0 to UBound(ALL_PANELS_ARRAY, 2)
                            If second_indicator = "" Then
                                If ALL_PANELS_ARRAY(the_panel_const, case_panels) = panel_indicator Then
                                    If ALL_PANELS_ARRAY(one_per_case_const, case_panels) = FALSE Then
                                        the_memb = ALL_PANELS_ARRAY(the_memb_const, case_panels)
                                        the_inst = ALL_PANELS_ARRAY(the_inst_const, case_panels)
                                        If the_memb = "" Then the_memb = "XX"
                                        If the_inst = "" Then the_inst = "XX"
                                        If panel_array_to_use = case_panels Then
                                            Text x_pos + 3, 342, 35, 10, the_memb & " - " & the_inst
                                        Else
                                            PushButton x_pos, 340, 35, 10, the_memb & " - " & the_inst, ALL_PANELS_ARRAY(panel_btn_const, case_panels)
                                        End If
                                        x_pos = x_pos + 35
                                    End If
                                End If
                            Else
                                If ALL_PANELS_ARRAY(the_panel_const, case_panels) = panel_indicator OR ALL_PANELS_ARRAY(the_panel_const, case_panels) = second_indicator Then
                                    the_memb = ALL_PANELS_ARRAY(the_memb_const, case_panels)
                                    the_inst = ALL_PANELS_ARRAY(the_inst_const, case_panels)
                                    If the_memb = "" Then the_memb = "XX"
                                    If the_inst = "" Then the_inst = "XX"
                                    If panel_array_to_use = case_panels Then
                                        Text x_pos + 3, 342, 50, 10, ALL_PANELS_ARRAY(the_panel_const, case_panels) & " "& the_memb & " - " & the_inst
                                    Else
                                        PushButton x_pos, 340, 50, 10, ALL_PANELS_ARRAY(the_panel_const, case_panels) & " "& the_memb & " - " & the_inst, ALL_PANELS_ARRAY(panel_btn_const, case_panels)
                                    End If
                                    x_pos = x_pos + 50
                                End If
                            End If
                        Next
                        ' PushButton x_pos, 340, 50, 10, "Add Another", add_another_btn


                        PushButton 380, 355, 50, 10, "PREVIOUS", previous_btn
                        PushButton 500, 350, 50, 15, "FINISH", finish_btn
                        CancelButton 550, 350, 50, 15
                        ' OkButton 550, 330, 50, 15
                      If panel_array_to_use <> "" Then
                          GroupBox 5, 5, 515, 330, ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use) & " - " & ALL_PANELS_ARRAY(the_memb_const, panel_array_to_use) & " - " & ALL_PANELS_ARRAY(the_inst_const, panel_array_to_use)
                      ElseIf panel_indicator = "NOTES" Then
                          GroupBox 5, 5, 515, 330, "NOTES on Case at CSR Processing"
                      ElseIf panel_indicator = "VERIFS" Then
                          GroupBox 5, 5, 515, 330, "VERIFICATIONS Needed"
                      Else
                          Text 5, 5, 515, 10, "There are no " & panel_indicator & " panels on this case."
                      End If
                      If panel_indicator = "ADDR" Then
                          CheckBox 20, 20, 130, 10, "Check here to have the script update ", update_addr_checkbox
                          Text 20, 30, 150, 10, "the ADDR Panel with this new information."
                          Text 20, 50, 95, 10, "Current Residence Address"
                          Text 30, 65, 115, 10, resi_line_one
                          If resi_line_two = "" Then
                             Text 30, 75, 115, 10, resi_city & ", " & resi_state & " " & resi_zip
                          Else
                              Text 30, 75, 115, 10, resi_line_two
                              Text 30, 85, 115, 10, resi_city & ", " & resi_state & " " & resi_zip
                          End If
                          If residence_address_match_yn = "No - there is a difference." Then
                              Text 400, 15, 50, 10, "Effective Date"
                              EditBox 460, 10, 50, 15, new_addr_effective_date
                              Text 170, 15, 145, 10, "New Residence Address Reported on CSR:"
                              Text 180, 35, 45, 10, "House/Street:"
                              EditBox 230, 30, 280, 15, new_resi_one
                              Text 210, 55, 15, 10, "City:"
                              EditBox 230, 50, 80, 15, new_resi_city
                              Text 320, 55, 20, 10, "State:"
                              DropListBox 345, 50, 75, 45, state_list, new_resi_state
                              Text 435, 55, 20, 10, "Zip:"
                              EditBox 460, 50, 50, 15, new_resi_zip
                              Text 200, 75, 30, 10, "County:"
                              DropListBox 230, 70, 190, 45, "Select One..."+chr(9)+county_list, new_resi_county
                              Text 255, 90, 90, 10, "Address/Home Verification:"
                              DropListBox 350, 85, 125, 45, "Select One..."+chr(9)+"SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Ver Prvd", new_shel_verif
                          End If

                          Text 255, 110, 60, 10, "Homeless Status:"
                          DropListBox 320, 105, 190, 45, "Select One..."+chr(9)+"Yes - Homeless"+chr(9)+"No", homeless_status
                          Text 255, 125, 60, 10, "Living Situation:"
                          DropListBox 320, 120, 190, 45, "Select One..."+chr(9)+"01 - Own home, lease or roomate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown", living_situation_status

                          Text 20, 145, 95, 10, "Current Mailing Address"
                          If mail_line_one = "" Then
                              Text 30, 160, 115, 10, "NO MAILING ADDRESS LISTED"
                          Else
                              Text 30, 160, 115, 10, mail_line_one
                              If mail_line_two = "" THen
                                  Text 30, 170, 115, 10, mail_city & ", " & mail_state & " " & mail_zip
                              Else
                                  Text 30, 170, 115, 10, mail_line_two
                                  Text 30, 180, 115, 10, mail_city & ", " & mail_state & " " & mail_zip
                              End If
                          End If
                          If mailing_address_match_yn = "No - there is a difference." Then
                              Text 170, 145, 145, 10, "New Mailing Address Reported on CSR:"
                              Text 180, 165, 45, 10, "House/Street:"
                              EditBox 230, 160, 280, 15, new_mail_one
                              Text 210, 185, 15, 10, "City:"
                              EditBox 230, 180, 80, 15, new_mail_city
                              Text 320, 185, 20, 10, "State:"
                              DropListBox 345, 180, 75, 45, state_list, new_mail_state
                              Text 435, 185, 20, 10, "Zip:"
                              EditBox 460, 180, 50, 15, new_mail_zip
                          End If

                          Text 20, 210, 95, 10, "Current Phone Numbers:"
                          Text 30, 225, 120, 10, curr_phone_one & "          " & curr_phone_type_one
                          Text 30, 240, 120, 10, curr_phone_two & "          " & curr_phone_type_two
                          Text 30, 255, 120, 10, curr_phone_three & "          " & curr_phone_type_three

                          Text 170, 210, 145, 10, "New Phone Number Reported on CSR:"
                          Text 180, 230, 170, 10, "Phone One:" & "                                                  " & "Type:"
                          EditBox 225, 225, 80, 15, new_phone_one
                          DropListBox 345, 225, 80, 45, "Select ..."+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"C - Cell"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_one_type

                          Text 180, 250, 170, 10, "Phone Two:" & "                                                  " & "Type:"
                          EditBox 225, 245, 80, 15, new_phone_two
                          DropListBox 345, 245, 80, 45, "Select ..."+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"C - Cell"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_two_type

                          Text 175, 270, 175, 10, "Phone Three:" & "                                                  " & "Type:"
                          EditBox 225, 265, 80, 15, new_phone_three
                          DropListBox 345, 265, 80, 45, "Select ..."+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"C - Cell"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_three_type

                          Text 10, 295, 50, 10, "Address Notes:"
                          EditBox 65, 290, 450, 15, notes_on_address
                          ' Text 360, 315, 45, 10, "Update Date:"
                          ' EditBox 410, 310, 60, 15, panel_update_date
                      ElseIf panel_indicator = "MEMB" Then
                          If panel_array_to_use <> "" Then
                              pers_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              ' MsgBox "Person to use - " & pers_to_use
                              GroupBox 15, 15, 265, 125, "MEMB"
                              Text 25, 30, 170, 10, ALL_CLIENTS_ARRAY(memb_first_name, pers_to_use) & " " & ALL_CLIENTS_ARRAY(memb_last_name, pers_to_use)
                              Text 125, 40, 150, 10, "Identity Verif: " & ALL_CLIENTS_ARRAY(memb_id_verif, pers_to_use)
                              Text 230, 30, 35, 10, "Age: " & ALL_CLIENTS_ARRAY(memb_age, pers_to_use)
                              Text 25, 55, 200, 10, "SSN: " & ALL_CLIENTS_ARRAY(memb_soc_sec_numb, pers_to_use) & "               Verif: " & ALL_CLIENTS_ARRAY(memb_ssn_verif, pers_to_use)
                              Text 25, 65, 200, 10, "DOB: " & ALL_CLIENTS_ARRAY(memb_dob, pers_to_use) & "               Verif: " & ALL_CLIENTS_ARRAY(memb_dob_verif, pers_to_use)
                              Text 25, 80, 300, 10, "Rel to Applicant: " & ALL_CLIENTS_ARRAY(memb_rel_to_applct, pers_to_use)
                              Text 25, 95, 150, 10, "Spoken Language: " & ALL_CLIENTS_ARRAY(memb_spoken_language, pers_to_use)
                              Text 25, 105, 150, 10, "Written Language: " & ALL_CLIENTS_ARRAY(memb_written_language, pers_to_use)
                              Text 230, 55, 35, 10, "Alias: " & ALL_CLIENTS_ARRAY(memb_alias, pers_to_use)
                              Text 185, 95, 75, 10, "Interpreter Needed: " & ALL_CLIENTS_ARRAY(memb_interpreter, pers_to_use)
                              Text 25, 120, 150, 10, "Race: " & ALL_CLIENTS_ARRAY(memb_race, pers_to_use)
                              Text 205, 120, 55, 10, "Hispanic/Lat.: " & ALL_CLIENTS_ARRAY(memb_ethnicity, pers_to_use)

                              GroupBox 15, 145, 265, 135, "MEMI"
                              Text 25, 155, 250, 10, "Marital Status: " & ALL_CLIENTS_ARRAY(memi_marriage_status, pers_to_use)
                              Text 25, 165, 250, 10, "Spouse: " & ALL_CLIENTS_ARRAY(memi_spouse_ref, pers_to_use) & " - " & ALL_CLIENTS_ARRAY(memi_spouse_name, pers_to_use)
                              GroupBox 20, 175, 255, 40, "Designated Spouse"
                              If ALL_CLIENTS_ARRAY(memi_designated_spouse, pers_to_use) <> "" Then
                                  Text 30, 185, 200, 10, "Designated Spouse: " & ALL_CLIENTS_ARRAY(memi_designated_spouse, pers_to_use)
                                  Text 30, 200, 250, 10, "Marriage Date: " & ALL_CLIENTS_ARRAY(memi_marriage_date, pers_to_use) & "     Marriage Date Verif: " & ALL_CLIENTS_ARRAY(memi_marriage_verif, pers_to_use)
                              End If
                              Text 25, 220, 250, 10, "Citizen: " & ALL_CLIENTS_ARRAY(memi_citizen, pers_to_use) & "                  Verif: " & ALL_CLIENTS_ARRAY(memi_citizen_verif, pers_to_use)
                              Text 25, 230, 250, 10, "Last Grade Completed: " & ALL_CLIENTS_ARRAY(memi_last_grade, pers_to_use)
                              Text 25, 245, 250, 10, "In MN>12 Months: " & ALL_CLIENTS_ARRAY(memi_MN_entry_date, pers_to_use) & "                       Residence Verif: " & ALL_CLIENTS_ARRAY(memi_resi_verif, pers_to_use)
                              Text 25, 255, 245, 10, "MN Entry Date: " & ALL_CLIENTS_ARRAY(memi_MN_entry_date, pers_to_use) & "                            Former State: " & ALL_CLIENTS_ARRAY(memi_former_state, pers_to_use)
                              Text 25, 265, 110, 10, "Other St FS End Date: " & ALL_CLIENTS_ARRAY(memi_other_FS_end, pers_to_use)

                              y_pos = 25
                              Text 285, 15, 200, 10, "Q2. Anyone move in or out - " & quest_two_move_in_out
                              If client_moved_in = FALSE Then
                                  Text 300, y_pos, 200, 10, "No one indicated as moving in."
                                  y_pos = y_pos + 10
                              Else
                                  Text 300, y_pos, 200, 10, "Clients moved in:"
                                  y_pos = y_pos + 10
                                  For known_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
                                      If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then Text 310, y_pos, 175, 10, ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & " - " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb) & ", " & ALL_CLIENTS_ARRAY(memb_first_name, known_memb)
                                      y_pos = y_pos + 10
                                  Next
                              End If

                              If client_moved_out = FALSE Then
                                  Text 300, y_pos, 200, 10, "No one indicated as moving out."
                                  y_pos = y_pos + 10
                              Else
                                  Text 300, y_pos, 200, 10, "Clients moved out:"
                                  y_pos = y_pos + 10
                                  For known_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
                                      If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then Text 310, y_pos, 175, 10, ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & " - " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb) & ", " & ALL_CLIENTS_ARRAY(memb_first_name, known_memb)
                                      y_pos = y_pos + 10
                                  Next
                              End If

                              Text 285, y_pos + 5, 35, 10, "Q2 Notes:"
                              EditBox 325, y_pos, 185, 15, question_two_notes
                              y_pos = y_pos + 40

                              Text 285, y_pos, 200, 10, "Q4. Anyone applying for MA - " & apply_for_ma
                              y_pos = y_pos + 10
                              Text 300, y_pos, 105, 10, "Members requesting MA:"
                              y_pos = y_pos + 10
                              If NEW_MA_REQUEST_ARRAY(ma_request_client, 0) = "Select or Type" Then
                                  Text 300, y_pos, 200, 10, "No members listed on form."
                                  y_pos = y_pos + 10
                              Else
                                  For each_new_memb = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
                                      Text 300, y_pos, 200, 10, NEW_MA_REQUEST_ARRAY(ma_request_client, each_new_memb)
                                      y_pos = y_pos + 10
                                  Next
                              End If
                              Text 285, y_pos + 5, 35, 10, "Q4 Notes:"
                              EditBox 325, y_pos, 185, 15, question_four_notes

                              Text 10, 295, 50, 10, "MEMB Notes:"
                              EditBox 65, 290, 450, 15, ALL_CLIENTS_ARRAY(memb_notes, ei_to_use)
                          End If
                      ElseIf panel_indicator = "WREG" Then
                          If panel_array_to_use <> "" Then
                              pers_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              Text 15, 25, 180, 10, "Household Member: " & ALL_CLIENTS_ARRAY(memb_first_name, pers_to_use) & " " & ALL_CLIENTS_ARRAY(memb_last_name, pers_to_use)
                              Text 15, 40, 175, 10, "FS PWE: " & ALL_CLIENTS_ARRAY(wreg_pwe, pers_to_use)
                              Text 15, 60, 250, 10, "FSET Work Reg Status: " & ALL_CLIENTS_ARRAY(wreg_status, pers_to_use)
                              Text 15, 70, 110, 10, "Defer FSET/No Funds: " & ALL_CLIENTS_ARRAY(wreg_defer_fset, pers_to_use)
                              Text 15, 80, 120, 10, "FSET Orientation Date: " & ALL_CLIENTS_ARRAY(wreg_fset_orient_date, pers_to_use)
                              Text 15, 95, 135, 10, "FSET Sanction Begin Date: " & ALL_CLIENTS_ARRAY(wreg_sanc_begin_date, pers_to_use)
                              Text 15, 105, 200, 10, "Number of Sanctions: " & ALL_CLIENTS_ARRAY(wreg_sanc_count, pers_to_use)
                              Text 15, 115, 200, 10, "Reason for Sanction: " & ALL_CLIENTS_ARRAY(wreg_sanc_reasons, pers_to_use)
                              Text 15, 130, 180, 10, "ABAWD Status: " & ALL_CLIENTS_ARRAY(wreg_abawd_status, pers_to_use)
                              Text 15, 140, 75, 10, "Banked Months: " & ALL_CLIENTS_ARRAY(wreg_banekd_months, pers_to_use)
                              Text 15, 155, 200, 10, "GA Eligibility Basis:" & ALL_CLIENTS_ARRAY(wreg_GA_basis, pers_to_use)
                              Text 15, 165, 200, 10, "GA Cooperation: " & ALL_CLIENTS_ARRAY(wreg_GA_coop, pers_to_use)

                              Text 15, 185, 200, 10, "Number of Counted ABAWD Months on WREG tracker: " & ALL_CLIENTS_ARRAY(wreg_numb_ABAWD_months, pers_to_use)
                              Text 15, 195, 200, 10, "List of Counted ABAWD Months: " & ALL_CLIENTS_ARRAY(wreg_ABAWD_months_list, pers_to_use)

                              Text 15, 215, 200, 10, "Number of Second Set Months on WREG Traker: " & ALL_CLIENTS_ARRAY(wreg_numb_second_set_months, pers_to_use)
                              Text 15, 225, 200, 10, "List of Second Set Months: " & ALL_CLIENTS_ARRAY(wreg_second_set_months_list, pers_to_use)

                              Text 285, 15, 200, 10, "Q19. Have you worked more than 20 hours per week - " & quest_nineteen_form_answer
                              Text 285, 30, 37, 10, "Q19 Notes:"
                              EditBox 325, 25, 185, 15, question_nineteen_notes

                              Text 10, 295, 50, 10, "WREG Notes:"
                              EditBox 65, 290, 450, 15, ALL_CLIENTS_ARRAY(wreg_notes, pers_to_use)
                          End If

                      ElseIf panel_indicator = "FACI" Then
                          If panel_array_to_use <> "" Then
                              faci_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              CheckBox 280, 25, 145, 10, "Check Here if verification for this FACI is needed.", FACILITIES_ARRAY(faci_verif_checkbox, faci_to_use)

                              Text 15, 20, 225, 10, "Member in FACI: Member " & FACILITIES_ARRAY(faci_ref_numb, faci_to_use) & " - " & FACILITIES_ARRAY(faci_member, faci_to_use)
                              Text 20, 45, 195, 10, "Facility Name: " & FACILITIES_ARRAY(faci_name, faci_to_use)
                              Text 20, 60, 100, 10, "Vandor Number: " & FACILITIES_ARRAY(faci_vendor_number, faci_to_use)
                              Text 20, 75, 145, 10, "Facility Type: " & FACILITIES_ARRAY(faci_type, faci_to_use)
                              Text 20, 90, 50, 10, "FS Elig: " & FACILITIES_ARRAY(faci_FS_elig, faci_to_use)
                              Text 85, 90, 175, 10, "FS FACI Type: " & FACILITIES_ARRAY(faci_FS_type, faci_to_use)
                              GroupBox 15, 35, 240, 70, "FACI Information"
                              Text 15, 110, 95, 10, "Waiver Type: " & FACILITIES_ARRAY(faci_waiver_type, faci_to_use)
                              Text 15, 125, 180, 10, "LTC Inelig Reason: " & FACILITIES_ARRAY(faci_ltc_inelig_reason, faci_to_use)
                              Text 15, 140, 240, 10, "LTC Pre-Screen/Inelig - Begin Date: " & FACILITIES_ARRAY(faci_inelig_begin_date, faci_to_use) & "       End Date: " & FACILITIES_ARRAY(faci_inelig_end_date, faci_to_use)
                              Text 15, 155, 130, 10, "Anticipated Date Out: " & FACILITIES_ARRAY(faci_anticipated_out_date, faci_to_use)
                              Text 15, 170, 150, 10, "GRH Plan    Required: " & FACILITIES_ARRAY(faci_GRH_plan_required, faci_to_use) & "        Verif: " & FACILITIES_ARRAY(faci_GRH_plan_verif, faci_to_use)
                              Text 15, 185, 250, 10, "County App Placement: " & FACILITIES_ARRAY(faci_cty_app_place, faci_to_use) & "          Approval Cty: " & FACILITIES_ARRAY(faci_approval_cty_name, faci_to_use)
                              Text 15, 200, 115, 10, "GRH DOC Amount $ " & FACILITIES_ARRAY(faci_GRH_DOC_amount, faci_to_use)
                              Text 145, 200, 50, 10, "Postpay: " & FACILITIES_ARRAY(faci_GRH_postpay, faci_to_use)
                              Text 15, 220, 40, 10, "GRH Rate"
                              Text 85, 220, 30, 10, "Date In"
                              Text 145, 220, 30, 10, "Date Out"
                              Text 20, 235, 50, 10, FACILITIES_ARRAY(faci_stay_one_rate, faci_to_use)
                              Text 80, 235, 150, 10, FACILITIES_ARRAY(faci_stay_one_date_in, faci_to_use) & "           " & FACILITIES_ARRAY(faci_stay_one_date_out, faci_to_use)
                              ' Text 80, 235, 35, 10, FACILITIES_ARRAY(faci_stay_one_date_in, faci_to_use)
                              ' Text 145, 235, 35, 10, FACILITIES_ARRAY(faci_stay_one_date_out, faci_to_use)
                              Text 20, 245, 50, 10, FACILITIES_ARRAY(faci_stay_two_rate, faci_to_use)
                              Text 80, 245, 150, 10, FACILITIES_ARRAY(faci_stay_two_date_in, faci_to_use) & "           " & FACILITIES_ARRAY(faci_stay_two_date_out, faci_to_use)
                              ' Text 80, 245, 35, 10, FACILITIES_ARRAY(faci_stay_two_date_in, faci_to_use)
                              ' Text 145, 245, 35, 10, FACILITIES_ARRAY(faci_stay_two_date_out, faci_to_use)
                              Text 20, 255, 50, 10, FACILITIES_ARRAY(faci_stay_three_rate, faci_to_use)
                              Text 80, 255, 150, 10, FACILITIES_ARRAY(faci_stay_three_date_in, faci_to_use) & "           " & FACILITIES_ARRAY(faci_stay_three_date_out, faci_to_use)
                              ' Text 80, 255, 35, 10, FACILITIES_ARRAY(faci_stay_three_date_in, faci_to_use)
                              ' Text 145, 255, 35, 10, FACILITIES_ARRAY(faci_stay_three_date_out, faci_to_use)
                              Text 20, 265, 50, 10, FACILITIES_ARRAY(faci_stay_four_rate, faci_to_use)
                              Text 80, 265, 150, 10, FACILITIES_ARRAY(faci_stay_four_date_in, faci_to_use) & "           " & FACILITIES_ARRAY(faci_stay_four_date_out, faci_to_use)
                              ' Text 80, 265, 35, 10, FACILITIES_ARRAY(faci_stay_four_date_in, faci_to_use)
                              ' Text 145, 265, 35, 10, FACILITIES_ARRAY(faci_stay_four_date_out, faci_to_use)
                              Text 20, 275, 50, 10, FACILITIES_ARRAY(faci_stay_five_rate, faci_to_use)
                              Text 80, 275, 150, 10, FACILITIES_ARRAY(faci_stay_five_date_in, faci_to_use) & "           " & FACILITIES_ARRAY(faci_stay_five_date_out, faci_to_use)
                              ' Text 80, 275, 35, 10, FACILITIES_ARRAY(faci_stay_five_date_in, faci_to_use)
                              ' Text 145, 275, 35, 10, FACILITIES_ARRAY(faci_stay_five_date_out, faci_to_use)

                              Text 10, 295, 50, 10, "FACI Notes:"
                              EditBox 65, 290, 450, 15, FACILITIES_ARRAY(faci_notes, faci_to_use)
                          End If
                      ElseIf panel_indicator = "REVW" Then

                          GroupBox 15, 20, 250, 175, "CSR Information (REVW)"
                          Text 25, 30, 200, 10, "CASH/GRH - "
                          Text 30, 40, 195, 10, "Review Status: " & panel_grh_status
                          Text 30, 50, 195, 10, "SR: Next Review: " & panel_grh_sr_month & "/" & panel_grh_sr_year
                          Text 30, 60, 195, 10, "ER: Next Review: " & panel_grh_er_month & "/" & panel_grh_er_year

                          Text 25, 80, 200, 10, "SNAP - "
                          Text 30, 90, 195, 10, "Review Status: " & panel_snap_status
                          Text 30, 100, 195, 10, "SR: Next Review: " & panel_snap_sr_month & "/" & panel_snap_sr_year
                          Text 30, 110, 195, 10, "ER: Next Review: " & panel_snap_er_month & "/" & panel_snap_er_year

                          Text 25, 130, 200, 10, "HC - "
                          Text 30, 140, 195, 10, "Review Status: " & panel_hc_status
                          Text 30, 150, 195, 10, "SR: Next Review: " & panel_hc_sr_month & "/" & panel_hc_sr_year
                          Text 30, 160, 195, 10, "ER: Next Review: " & panel_hc_er_month & "/" & panel_hc_er_year

                          GroupBox 15, 200, 250, 70, "HRF Information (MONT)"
                          Text 25, 215, 100, 40, mont_detail
                      ElseIf panel_indicator = "ACCT" Then
                          If panel_array_to_use <> "" Then
                              asset_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              If ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use) = "CASH" Then
                                  Text 20, 25, 170, 10, "Owner: " &  ALL_ASSETS_ARRAY(owner_name, asset_to_use)
                                  Text 20, 40, 165, 10, "CASH Asset"
                                  Text 20, 55, 100, 10, "Amount: $ " &  ALL_ASSETS_ARRAY(amount_const, asset_to_use)
                                  y_pos = 25
                              Else
                                  CheckBox 280, 25, 145, 10, "Check Here if ACCT needs to be Verified.", ALL_ASSETS_ARRAY(verif_checkbox_const, asset_to_use)
                                  Text 345, 40, 75, 10, "Time period to verify:"
                                  EditBox 425, 35, 85, 15, ALL_ASSETS_ARRAY(verif_time_const, asset_to_use)

                                  Text 20, 25, 170, 10, "Owner: " &  ALL_ASSETS_ARRAY(owner_name, asset_to_use)
                                  Text 20, 40, 165, 10, "Account Type: " &  ALL_ASSETS_ARRAY(type_const, asset_to_use)
                                  Text 20, 50, 165, 10, "Account Number: " &  ALL_ASSETS_ARRAY(id_number_const, asset_to_use)
                                  Text 20, 60, 165, 10, "Account Location: " &  ALL_ASSETS_ARRAY(name_const, asset_to_use)
                                  Text 20, 80, 100, 10, "Balance: $ " &  ALL_ASSETS_ARRAY(amount_const, asset_to_use)
                                  Text 20, 90, 130, 10, "Verification: " &  ALL_ASSETS_ARRAY(verif_const, asset_to_use)
                                  Text 20, 100, 75, 10, "As Of: " & ALL_ASSETS_ARRAY(balance_date_const, asset_to_use)
                                  Text 20, 115, 120, 10, "Withdrawl Penalty: $ " & ALL_ASSETS_ARRAY(withdraw_penalty_const, asset_to_use)
                                  Text 20, 125, 50, 10, "Withdraw: " & ALL_ASSETS_ARRAY(withdraw_yn_const, asset_to_use)
                                  Text 20, 135, 130, 10, "Verification: " & ALL_ASSETS_ARRAY(withdraw_verif_const, asset_to_use)
                                  GroupBox 20, 155, 190, 30, "Count (Y/N)"
                                  Text 30, 170, 30, 10, "Cash: " & ALL_ASSETS_ARRAY(count_cash_const, asset_to_use)
                                  Text 70, 170, 30, 10, "SNAP: " & ALL_ASSETS_ARRAY(count_snap_const, asset_to_use)
                                  Text 110, 170, 25, 10, "HC: " & ALL_ASSETS_ARRAY(count_hc_const, asset_to_use)
                                  Text 145, 170, 25, 10, "GRH: " & ALL_ASSETS_ARRAY(count_grh_const, asset_to_use)
                                  Text 180, 170, 25, 10, "IV-E: " & ALL_ASSETS_ARRAY(count_ive_const, asset_to_use)
                                  Text 20, 195, 60, 10, "Joint Owner: " & ALL_ASSETS_ARRAY(joint_own_const, asset_to_use)
                                  Text 115, 195, 60, 10, "Share Ratio: " & ALL_ASSETS_ARRAY(share_ratio_const, asset_to_use)
                                  Text 20, 210, 110, 10, "Next Interest Date: " & ALL_ASSETS_ARRAY(next_interst_const, asset_to_use)

                                  Text 10, 295, 50, 10, "ACCT Notes:"
                                  EditBox 65, 290, 450, 15, ALL_ASSETS_ARRAY(item_notes_const, asset_to_use)
                                  y_pos = 55
                              End If

                              Text 285, y_pos, 200, 10, "Q9. Anyone have liquid assets - " & ma_liquid_assets
                              y_pos = y_pos + 10

                              first_account = TRUE
                              For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
                                If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
                                  If first_account = TRUE Then
                                      ' Text 300, y_pos, 55, 10, "Owner(s) Name"
                                      ' Text 360, y_pos, 25, 10, "Type"
                                      ' Text 390, y_pos, 50, 10, "Bank Name"
                                      Text 300, y_pos, 200, 10, "Owner Name   &    Type    &    Bank Name"
                                      y_pos = y_pos + 10
                                      first_account = FALSE
                                  End If
                                  Text 300, y_pos, 200, 10, NEW_ASSET_ARRAY(asset_client, each_asset) & " - " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " - " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
                                  ' Text 300, y_pos, 130, 45, NEW_ASSET_ARRAY(asset_client, each_asset) 'liquid_asset_member'
                                  ' Text 360, y_pos, 115, 40, NEW_ASSET_ARRAY(asset_acct_type, each_asset)'liquid_asst_type
                                  ' Text 390, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)'liquid_asset_name
                                  y_pos = y_pos + 10
                                End If
                              Next
                              If first_account = TRUE Then
                                Text 300, y_pos, 200, 10, "No assets were added."
                                y_pos = y_pos + 10
                              End If

                              Text 285, y_pos + 5, 35, 10, "Q9 Notes:"
                              EditBox 325, y_pos, 185, 15, question_nine_notes

                          End If
                      ElseIf panel_indicator = "SECU" Then
                          If panel_array_to_use <> "" Then
                              asset_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)

                              CheckBox 280, 25, 145, 10, "Check Here if SECU needs to be Verified.", ALL_ASSETS_ARRAY(verif_checkbox_const, asset_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_ASSETS_ARRAY(verif_time_const, asset_to_use)

                              Text 20, 25, 170, 10, "Owner: " &  ALL_ASSETS_ARRAY(owner_name, asset_to_use)
                              Text 20, 40, 165, 10, "Security Type: " &  ALL_ASSETS_ARRAY(type_const, asset_to_use)
                              Text 20, 50, 165, 10, "Account Number: " &  ALL_ASSETS_ARRAY(id_number_const, asset_to_use)
                              Text 20, 60, 165, 10, "Security Name: " &  ALL_ASSETS_ARRAY(name_const, asset_to_use)
                              Text 20, 80, 135, 10, "Cash Surrender Value: $ " &  ALL_ASSETS_ARRAY(amount_const, asset_to_use)
                              Text 20, 90, 130, 10, "Verification: " &  ALL_ASSETS_ARRAY(verif_const, asset_to_use)
                              Text 20, 100, 75, 10, "As Of: " & ALL_ASSETS_ARRAY(balance_date_const, asset_to_use)
                              Text 20, 110, 135, 10, "Face Value: $ " & ALL_ASSETS_ARRAY(face_value_const, asset_to_use)
                              Text 20, 130, 120, 10, "Withdrawl Penalty: $ " & ALL_ASSETS_ARRAY(withdraw_penalty_const, asset_to_use)
                              Text 20, 140, 50, 10, "Withdraw: " & ALL_ASSETS_ARRAY(withdraw_yn_const, asset_to_use)
                              Text 20, 150, 130, 10, "Verification: " & ALL_ASSETS_ARRAY(withdraw_verif_const, asset_to_use)
                              GroupBox 20, 170, 190, 30, "Count (Y/N)"
                              Text 30, 185, 30, 10, "Cash: " & ALL_ASSETS_ARRAY(count_cash_const, asset_to_use)
                              Text 70, 185, 30, 10, "SNAP: " & ALL_ASSETS_ARRAY(count_snap_const, asset_to_use)
                              Text 110, 185, 25, 10, "HC: " & ALL_ASSETS_ARRAY(count_hc_const, asset_to_use)
                              Text 145, 185, 25, 10, "GRH: " & ALL_ASSETS_ARRAY(count_grh_const, asset_to_use)
                              Text 180, 185, 25, 10, "IV-E: " & ALL_ASSETS_ARRAY(count_ive_const, asset_to_use)
                              Text 20, 210, 60, 10, "Joint Owner: " & ALL_ASSETS_ARRAY(joint_own_const, asset_to_use)
                              Text 115, 210, 60, 10, "Share Ratio: " & ALL_ASSETS_ARRAY(share_ratio_const, asset_to_use)

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q10. Anyone have securities - " & ma_security_assets
                              y_pos = y_pos + 10
                              first_secu = TRUE
                              For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
                                If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
                                    If first_secu = TRUE Then
                                        ' Text 300, y_pos, 55, 10, "Owner(s) Name"
                                        ' Text 360, y_pos, 25, 10, "Type"
                                        ' Text 390, y_pos, 50, 10, "Bank Name"
                                        Text 300, y_pos, 200, 10, "Owner Name   &    Type    &    Bank Name"
                                        y_pos = y_pos + 10
                                        first_secu = FALSE
                                    End If
                                    Text 300, y_pos, 200, 10, NEW_ASSET_ARRAY(asset_client, each_asset) & " - " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " - " & NEW_ASSET_ARRAY(asset_bank_name, each_asset)
                                    ' Text 300, y_pos, 130, 45, NEW_ASSET_ARRAY(asset_client, each_asset) 'security_asset_member
                                    ' Text 360, y_pos, 115, 40, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'security_asset_type
                                    ' Text 390, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_bank_name, each_asset)   'security_asset_name
                                    y_pos = y_pos + 10
                                End If
                              Next
                              If first_secu = TRUE Then
                                  Text 30, y_pos, 200, 10, "No assets added."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 38, 10, "Q10 Notes:"
                              EditBox 325, y_pos, 185, 15, question_ten_notes

                              Text 10, 295, 50, 10, "SECU Notes:"
                              EditBox 65, 290, 450, 15, ALL_ASSETS_ARRAY(item_notes_const, asset_to_use)
                          End If
                      ElseIf panel_indicator = "CARS" Then
                          If panel_array_to_use <> "" Then
                              asset_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)

                              CheckBox 280, 25, 145, 10, "Check Here if CARS needs to be Verified.", ALL_ASSETS_ARRAY(verif_checkbox_const, asset_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_ASSETS_ARRAY(verif_time_const, asset_to_use)

                              Text 20, 25, 170, 10, "Owner: " &  ALL_ASSETS_ARRAY(owner_name, asset_to_use)
                              Text 20, 40, 165, 10, "Vehicle Type: " &  ALL_ASSETS_ARRAY(type_const, asset_to_use)
                              Text 20, 55, 60, 10, "Year: " &  ALL_ASSETS_ARRAY(year_const, asset_to_use)
                              Text 20, 65, 100, 10, "Make: " &  ALL_ASSETS_ARRAY(make_const, asset_to_use)
                              Text 20, 75, 100, 10, "Model: " &  ALL_ASSETS_ARRAY(model_const, asset_to_use)
                              Text 20, 90, 110, 10, "Trade-In Value: $ " & ALL_ASSETS_ARRAY(trade_in_const, asset_to_use)
                              Text 20, 100, 110, 10, "Loan Value: $ " & ALL_ASSETS_ARRAY(loan_const, asset_to_use)
                              Text 20, 110, 135, 10, "Source: " & ALL_ASSETS_ARRAY(source_const, asset_to_use)
                              Text 20, 125, 135, 10, "Verification: " &  ALL_ASSETS_ARRAY(verif_const, asset_to_use)
                              Text 20, 140, 100, 10, "Amount Owed: $ " & ALL_ASSETS_ARRAY(owed_amt_const, asset_to_use)
                              Text 20, 150, 130, 10, "Verification: " & ALL_ASSETS_ARRAY(owed_verif_const, asset_to_use)
                              Text 20, 160, 80, 10, "As Of: " & ALL_ASSETS_ARRAY(owed_date_const, asset_to_use)
                              Text 20, 175, 150, 10, "Use: " & ALL_ASSETS_ARRAY(cars_use_const, asset_to_use)
                              Text 20, 185, 95, 10, "HC Client Benefit: " & ALL_ASSETS_ARRAY(hc_benefit_const, asset_to_use)
                              Text 20, 200, 60, 10, "Joint Owner: " & ALL_ASSETS_ARRAY(joint_own_const, asset_to_use)
                              Text 115, 200, 100, 10, "Share Ratio: " & ALL_ASSETS_ARRAY(share_ratio_const, asset_to_use)

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q11. Anyone have a vehicle - " & ma_vehicle
                              y_pos = y_pos + 10
                              first_car = TRUE
                              For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
                                  If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
                                      If first_car = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Owner(s) Name"
                                          ' Text 360, y_pos, 25, 10, "Type"
                                          ' Text 390, y_pos, 75, 10, "Year/Make/Model"
                                          Text 300, y_pos, 200, 10, "Owner Name   &    Type   &    Year/Make/Model"
                                          y_pos = y_pos + 10
                                          first_car = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_ASSET_ARRAY(asset_client, each_asset) & " - " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " - " & NEW_ASSET_ARRAY(asset_year_make_model, each_asset)
                                      ' Text 300, y_pos, 130, 45, NEW_ASSET_ARRAY(asset_client, each_asset)     'vehicle_asset_member
                                      ' Text 360, y_pos, 115, 40, NEW_ASSET_ARRAY(asset_acct_type, each_asset)    'vehicle_asset_type
                                      ' Text 390, y_pos, 195, 15, NEW_ASSET_ARRAY(asset_year_make_model, each_asset)  'vehicle_asset_name
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_car = TRUE Then
                                Text 30, y_pos, 200, 10, "No CARS Information entered."
                                y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 38, 10, "Q11 Notes:"
                              EditBox 325, y_pos, 185, 15, question_eleven_notes

                              Text 10, 295, 50, 10, "CARS Notes:"
                              EditBox 65, 290, 450, 15, ALL_ASSETS_ARRAY(item_notes_const, asset_to_use)
                          End If
                      ElseIf panel_indicator = "REST" Then
                          If panel_array_to_use <> "" Then
                              asset_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)

                              CheckBox 280, 25, 145, 10, "Check Here if REST needs to be Verified.", ALL_ASSETS_ARRAY(verif_checkbox_const, asset_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_ASSETS_ARRAY(verif_time_const, asset_to_use)

                              Text 20, 25, 170, 10, "Owner: " &  ALL_ASSETS_ARRAY(owner_name, asset_to_use)
                              Text 20, 40, 165, 10, "Property Type: " &  ALL_ASSETS_ARRAY(type_const, asset_to_use)
                              Text 20, 55, 165, 10, "Verification: " &  ALL_ASSETS_ARRAY(verif_const, asset_to_use)
                              Text 20, 70, 95, 10, "Market Value: $ " & ALL_ASSETS_ARRAY(market_value_const, asset_to_use)
                              Text 20, 80, 145, 10, "Verification: " & ALL_ASSETS_ARRAY(value_verif_const, asset_to_use)
                              Text 20, 95, 95, 10, "Amount Owed: $ " & ALL_ASSETS_ARRAY(owed_amt_const, asset_to_use)
                              Text 20, 105, 140, 10, "Verification: " & ALL_ASSETS_ARRAY(owed_verif_const, asset_to_use)
                              Text 20, 115, 70, 10, "As Of: " & ALL_ASSETS_ARRAY(owed_date_const, asset_to_use)
                              Text 20, 130, 160, 10, "Property Status: " & ALL_ASSETS_ARRAY(rest_prop_status_const, asset_to_use)
                              Text 20, 145, 60, 10, "Joint Owner: " & ALL_ASSETS_ARRAY(joint_own_const, asset_to_use)
                              Text 115, 145, 70, 10, "Share Ratio: " & ALL_ASSETS_ARRAY(share_ratio_const, asset_to_use)
                              Text 20, 165, 150, 10, "IV-E Repayment Agreement Date: " & ALL_ASSETS_ARRAY(rest_repymt_date_const, asset_to_use)

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q12. Anyone have real estate - " & ma_real_assets
                              y_pos = y_pos + 10
                              first_home = TRUE
                              For each_asset = 0 to Ubound(NEW_ASSET_ARRAY, 2)
                                  If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
                                      If first_home = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Owner(s) Name"
                                          ' Text 360, y_pos, 25, 10, "Address"
                                          ' Text 390, y_pos, 75, 10, "Type of Property"
                                          Text 300, y_pos, 200, 10, "Owner Name    &    Address    &    Type of Property"
                                          y_pos = y_pos + 10
                                          first_home = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_ASSET_ARRAY(asset_client, each_asset) & " - " & NEW_ASSET_ARRAY(asset_address, each_asset) & " - " & NEW_ASSET_ARRAY(asset_acct_type, each_asset)
                                      ' Text 300, y_pos, 130, 45, NEW_ASSET_ARRAY(asset_client, each_asset)     'property_asset_member
                                      ' Text 360, y_pos, 150, 15, NEW_ASSET_ARRAY(asset_address, each_asset)      'property_asset_address
                                      ' Text 390, y_pos, 150, 40, NEW_ASSET_ARRAY(asset_acct_type, each_asset)     'property_asset_type
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_home = TRUE Then
                                Text 300, y_pos, 200, 10, "CSR form for Question 12 is listed as 'No' and no REST information has been added."
                                y_pos = y_pos + 10
                              End If
                              Text 285, 95, 38, 10, "Q12 Notes:"
                              EditBox 325, 90, 185, 15, question_twelve_notes

                              Text 10, 295, 50, 10, "REST Notes:"
                              EditBox 65, 290, 450, 15, ALL_ASSETS_ARRAY(item_notes_const, asset_to_use)
                          End If
                      ElseIf panel_indicator = "BUSI" Then
                          If panel_array_to_use <> "" Then
                              ei_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              CheckBox 280, 25, 145, 10, "Check Here if BUSI needs to be Verified.", ALL_INCOME_ARRAY(verif_checkbox_const, ei_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_INCOME_ARRAY(verif_time_const, ei_to_use)
                              Text 20, 20, 200, 10, "Household Member: " & ALL_INCOME_ARRAY(owner_name, ei_to_use)
                              Text 20, 35, 195, 10, "Self Employment Type: " & ALL_INCOME_ARRAY(type_const, ei_to_use)
                              Text 30, 45, 100, 10, "Income Start Date: " & ALL_INCOME_ARRAY(start_date_const, ei_to_use)
                              GroupBox 15, 60, 330, 190, "Income Calculation"
                              Text 100, 70, 175, 10, "Retro                     Prosp                       Verif"
                              ' Text 175, 70, 50, 10, "Prospective"
                              ' Text 60, 70, 175, 10, "Income          Expense             Verification"
                              ' Text 90, 80, 15, 10, "Exp"
                              ' Text 125, 80, 15, 10, "Verif"
                              ' Text 170, 80, 15, 10, "Inc"
                              ' Text 200, 80, 15, 10, "Exp"
                              ' Text 235, 80, 15, 10, "Verif"
                              Text 30, 85, 70, 10, "CASH    Income -"
                              Text 100, 85, 220, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_gross_retro, ei_to_use) & "              $ " & ALL_INCOME_ARRAY(busi_cash_gross_prosp, ei_to_use) & "               " & ALL_INCOME_ARRAY(busi_cash_income_verif, ei_to_use)
                              ' Text 100, 85, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_gross_retro, ei_to_use)
                              ' Text 160, 85, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_gross_prosp, ei_to_use)
                              ' Text 220, 85, 60, 10, ALL_INCOME_ARRAY(busi_cash_income_verif, ei_to_use)
                              Text 50, 95, 35, 10, "Expenses -"
                              Text 100, 95, 220, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_expense_retro, ei_to_use) & "              $ " & ALL_INCOME_ARRAY(busi_cash_expense_prosp, ei_to_use) & "               " & ALL_INCOME_ARRAY(busi_cash_expense_verif, ei_to_use)
                              ' Text 100, 95, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_expense_retro, ei_to_use)
                              ' Text 160, 95, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_expense_prosp, ei_to_use)
                              ' Text 220, 95, 60, 10, ALL_INCOME_ARRAY(busi_cash_expense_verif, ei_to_use)
                              Text 65, 105, 20, 10, "Net - "
                              Text 100, 105, 220, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_net_retro, ei_to_use) & "               $ " & ALL_INCOME_ARRAY(busi_cash_net_prosp, ei_to_use)
                              ' Text 100, 105, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_net_retro, ei_to_use)
                              ' Text 160, 105, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_cash_net_prosp, ei_to_use)

                              Text 30, 120, 70, 10, "SNAP    Income -"
                              Text 100, 120, 220, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_gross_retro, ei_to_use) & "              $ " & ALL_INCOME_ARRAY(busi_snap_gross_prosp, ei_to_use) & "               " & ALL_INCOME_ARRAY(busi_snap_income_verif, ei_to_use)
                              ' Text 100, 120, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_gross_retro, ei_to_use)
                              ' Text 160, 120, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_gross_prosp, ei_to_use)
                              ' Text 220, 120, 60, 10, ALL_INCOME_ARRAY(busi_snap_income_verif, ei_to_use)
                              Text 50, 130, 35, 10, "Expenses -"
                              Text 100, 130, 220, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_expense_retro, ei_to_use) & "              $ " & ALL_INCOME_ARRAY(busi_snap_expense_prosp, ei_to_use) & "               " & ALL_INCOME_ARRAY(busi_snap_expense_verif, ei_to_use)
                              ' Text 100, 130, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_expense_retro, ei_to_use)
                              ' Text 160, 130, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_expense_prosp, ei_to_use)
                              ' Text 220, 130, 60, 10, ALL_INCOME_ARRAY(busi_snap_expense_verif, ei_to_use)
                              Text 70, 140, 20, 10, "Net - "
                              Text 100, 140, 220, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_net_retro, ei_to_use) & "               $" & ALL_INCOME_ARRAY(busi_snap_net_prosp, ei_to_use)
                              ' Text 100, 140, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_net_retro, ei_to_use)
                              ' Text 160, 140, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_snap_net_prosp, ei_to_use)

                              Text 30, 155, 70, 10, "HC-A     Income -"
                              Text 160, 155, 160, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_a_gross_prosp, ei_to_use) & "                 " & ALL_INCOME_ARRAY(busi_hc_a_income_verif, ei_to_use)
                              ' Text 160, 155, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_a_gross_prosp, ei_to_use)
                              ' Text 220, 155, 60, 10, ALL_INCOME_ARRAY(busi_hc_a_income_verif, ei_to_use)
                              Text 50, 165, 35, 10, "Expenses -"
                              Text 160, 165, 160, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_a_expense_prosp, ei_to_use) & "                 " & ALL_INCOME_ARRAY(busi_hc_a_expense_verif, ei_to_use)
                              ' Text 160, 165, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_a_expense_prosp, ei_to_use)
                              ' Text 220, 165, 60, 10, ALL_INCOME_ARRAY(busi_hc_a_expense_verif, ei_to_use)
                              Text 70, 175, 20, 10, "Net - "
                              Text 160, 175, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_a_net_prosp, ei_to_use)

                              Text 30, 190, 70, 10, "HC-B     Income -"
                              Text 160, 190, 160, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_b_gross_prosp, ei_to_use) & "                 " & ALL_INCOME_ARRAY(busi_hc_b_income_verif, ei_to_use)
                              ' Text 160, 155, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_b_gross_prosp, ei_to_use)
                              ' Text 220, 155, 60, 10, ALL_INCOME_ARRAY(busi_hc_b_income_verif, ei_to_use)
                              Text 50, 200, 35, 10, "Expenses -"
                              Text 160, 200, 160, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_b_expense_prosp, ei_to_use) & "                 " & ALL_INCOME_ARRAY(busi_hc_b_expense_verif, ei_to_use)
                              ' Text 160, 165, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_b_expense_prosp, ei_to_use)
                              ' Text 220, 165, 60, 10, ALL_INCOME_ARRAY(busi_hc_b_expense_verif, ei_to_use)
                              Text 70, 210, 20, 10, "Net - "
                              Text 160, 210, 30, 10, "$ " & ALL_INCOME_ARRAY(busi_hc_b_net_prosp, ei_to_use)

                              Text 50, 225, 35, 10, "Reptd Hrs:"
                              Text 100, 225, 220, 10, ALL_INCOME_ARRAY(rptd_hours_const, ei_to_use) & "                              " & ALL_INCOME_ARRAY(rptd_hours_const, ei_to_use)
                              ' Text 100, 225, 15, 10, ALL_INCOME_ARRAY(rptd_hours_const, ei_to_use)
                              ' Text 175, 225, 15, 10, ALL_INCOME_ARRAY(rptd_hours_const, ei_to_use)
                              Text 45, 235, 45, 10, "Min Wg Hrs:"
                              Text 100, 235, 220, 10, ALL_INCOME_ARRAY(min_wg_hours_const, ei_to_use) & "                              " & ALL_INCOME_ARRAY(min_wg_hours_const, ei_to_use)
                              ' Text 100, 235, 15, 10, ALL_INCOME_ARRAY(min_wg_hours_const, ei_to_use)
                              ' Text 175, 235, 15, 10, ALL_INCOME_ARRAY(min_wg_hours_const, ei_to_use)
                              Text 15, 255, 250, 10, "Self Employment Method: " & ALL_INCOME_ARRAY(busi_se_method, ei_to_use) & "         Selection Date: " & ALL_INCOME_ARRAY(busi_se_method_date, ei_to_use)

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q5. Anyone in the household self-employed - " & ma_self_employed
                              y_pos = y_pos + 10
                              first_busi= TRUE
                              For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
                                  If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
                                      If first_busi = TRUE then
                                          ' Text 300, y_pos, 75, 10, "Name"
                                          ' Text 380, y_pos, 55, 10, "Business Name"
                                          ' Text 440, y_pos, 35, 10, "Start Date"
                                          ' Text 480, y_pos, 50, 10, "Yearly Income"
                                          Text 300, y_pos, 200, 10, "Name    &    Business Name    &    Start Date    &    Income"
                                          y_pos = y_pos + 10
                                          first_busi = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_EARNED_ARRAY(earned_client, each_busi) & " - " &  NEW_EARNED_ARRAY(earned_source, each_busi) & " - " & NEW_EARNED_ARRAY(earned_start_date, each_busi) & " - " & NEW_EARNED_ARRAY(earned_amount, each_busi)
                                      ' Text 300, y_pos, 75, 45, NEW_EARNED_ARRAY(earned_client, each_busi)
                                      ' Text 380, y_pos, 55, 15, NEW_EARNED_ARRAY(earned_source, each_busi)
                                      ' Text 440, y_pos, 35, 15, NEW_EARNED_ARRAY(earned_start_date, each_busi)
                                      ' Text 480, y_pos, 50, 15, NEW_EARNED_ARRAY(earned_amount, each_busi)
                                      y_pos = y_pos  + 10
                                  End If
                              Next
                              If first_busi = TRUE Then
                                  Text 300, y_pos, 200, 10, "No BUSI Information entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 35, 10, "Q5 Notes:"
                              EditBox 325, y_pos, 185, 15, question_five_notes
                              y_pos = y_pos + 40

                              Text 285, y_pos, 200, 10, "Q16. Has there been a change in work income - " & quest_sixteen_form_answer
                              y_pos = y_pos + 10
                              first_income = TRUE
                              For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
                                  If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
                                      If first_income = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Client"
                                          ' Text 360, y_pos, 55, 10, "Employer (or Business Name)"
                                          ' Text 420, y_pos, 30, 10, "Change"
                                          ' Text 455, y_pos, 35, 10, "Amount"
                                          Text 200, y_pos, 200, 10, "Client    &    Employer    &    Change    &    Amount"
                                          y_pos = y_pos + 10
                                          first_income = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_EARNED_ARRAY(earned_client, the_earned) & " - " & NEW_EARNED_ARRAY(earned_source, the_earned) & " - " & NEW_EARNED_ARRAY(earned_change_date, the_earned) & " - " & NEW_EARNED_ARRAY(earned_amount, the_earned)
                                      ' Text 300, y_pos, 55, 45, NEW_EARNED_ARRAY(earned_client, the_earned)
                                      ' Text 360, y_pos, 55, 15, NEW_EARNED_ARRAY(earned_source, the_earned)
                                      ' Text 420, y_pos, 30, 15, NEW_EARNED_ARRAY(earned_change_date, the_earned)
                                      ' Text 455, y_pos, 35, 15, NEW_EARNED_ARRAY(earned_amount, the_earned)
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_income = TRUE Then
                                  Text 300, y_pos, 200, 10, "No income entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 35, 10, "Q16. Notes:"
                              EditBox 325, y_pos, 185, 15, question_sixteen_notes

                              Text 10, 295, 50, 10, "BUSI Notes:"
                              EditBox 65, 290, 450, 15, ALL_INCOME_ARRAY(item_notes_const, ei_to_use)
                          End If
                      ElseIf panel_indicator = "JOBS" Then
                          If panel_array_to_use <> "" Then
                              ei_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              Text 20, 25, 145, 10, "Employer: " & ALL_INCOME_ARRAY(name_const, ei_to_use)
                              Text 25, 40, 110, 10, "Income Type: " & ALL_INCOME_ARRAY(type_const, ei_to_use)
                              Text 130, 40, 120, 10, "Verification: " & ALL_INCOME_ARRAY(verif_const, ei_to_use)
                              Text 25, 50, 150, 10, "Subsidy: " & ALL_INCOME_ARRAY(jobs_subsidy, ei_to_use)
                              Text 185, 25, 80, 10, "Hourly Wage: $ " & ALL_INCOME_ARRAY(jobs_hrly_wage, ei_to_use)
                              Text 25, 65, 175, 10, "Pay Frequency: " & ALL_INCOME_ARRAY(frequency_const, ei_to_use)
                              Text 20, 80, 90, 10, "Income Start: " & ALL_INCOME_ARRAY(start_date_const, ei_to_use)
                              Text 130, 80, 80, 10, "Income End: " & ALL_INCOME_ARRAY(end_date_const, ei_to_use)
                              Text 20, 95, 235, 10, "Retrospective Income                               Prospective Income"
                              Text 20, 110, 100, 10, "Total Income:   $" & ALL_INCOME_ARRAY(retro_income_amount, ei_to_use)
                              Text 40, 120, 65, 10, "Hours:    " & ALL_INCOME_ARRAY(retro_income_hours, ei_to_use)
                              Text 150, 110, 90, 10, "Total Income:   $ " & ALL_INCOME_ARRAY(pay_amt_const, ei_to_use)
                              Text 170, 120, 65, 10, "Hours:    " & ALL_INCOME_ARRAY(pay_amt_const, ei_to_use)
                              GroupBox 15, 135, 200, 70, "SNAP PIC"
                              Text 25, 150, 115, 10, "Pay Frequency: " & ALL_INCOME_ARRAY(snap_pic_frequency, ei_to_use)
                              Text 25, 165, 135, 10, "Ave Hrs per Pay Date: " & ALL_INCOME_ARRAY(snap_pic_hours_per_pay, ei_to_use)
                              Text 25, 175, 140, 10, "Ave Income per Pay Date: $ " & ALL_INCOME_ARRAY(snap_pic_income_per_pay, ei_to_use)
                              Text 25, 190, 145, 10, "Prospective Monthly Income: $ " & ALL_INCOME_ARRAY(snap_pic_monthly_income, ei_to_use)
                              GroupBox 15, 210, 200, 60, "GRH PIC"
                              Text 25, 225, 115, 10, "Pay Frequency: " & ALL_INCOME_ARRAY(grh_pic_frequency, ei_to_use)
                              Text 25, 240, 140, 10, "Ave Income per Pay Date: $ " & ALL_INCOME_ARRAY(grh_pic_income_per_pay, ei_to_use)
                              Text 25, 255, 145, 10, "Prospective Monthly Income: $ " & ALL_INCOME_ARRAY(grh_pic_monthly_income, ei_to_use)
                              CheckBox 280, 25, 145, 10, "Check Here if JOB needs to be Verified.", ALL_INCOME_ARRAY(verif_checkbox_const, ei_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_INCOME_ARRAY(verif_time_const, ei_to_use)

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q6. Anyone work - " & ma_start_working
                              y_pos = y_pos + 10
                              first_job = TRUE
                              For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
                                  If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
                                      If first_job = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Name"
                                          ' Text 360, y_pos, 55, 10, "Employer Name"
                                          ' Text 420, y_pos, 30, 10, "Amount"
                                          ' Text 455, y_pos, 50, 10, "How often?"
                                          Text 300, y_pos, 200, 10, "Name    &    Employer    &    Amount    &    Frequency"
                                          y_pos = y_pos  + 10
                                          first_job = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_EARNED_ARRAY(earned_client, each_job) & " - " & NEW_EARNED_ARRAY(earned_source, each_job) & " - " & NEW_EARNED_ARRAY(earned_amount, each_job) & " - " & NEW_EARNED_ARRAY(earned_freq, each_job)
                                      ' Text 300, y_pos, 55, 45, NEW_EARNED_ARRAY(earned_client, each_job)
                                      ' Text 360, y_pos, 55, 15, NEW_EARNED_ARRAY(earned_source, each_job)
                                      ' Text 420, y_pos, 30, 15, NEW_EARNED_ARRAY(earned_amount, each_job)
                                      ' Text 455, y_pos, 50, 45, NEW_EARNED_ARRAY(earned_freq, each_job)
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_job = TRUE Then
                                  Text 300, y_pos, 200, 10, "No JOBS entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 35, 10, "Q6 Notes:"
                              EditBox 325, y_pos, 185, 15, question_six_notes
                              y_pos = y_pos + 40

                              Text 285, y_pos, 200, 10, "Q16. Has there been a change in work income - " & quest_sixteen_form_answer
                              y_pos = y_pos + 10
                              first_income = TRUE
                              For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
                                  If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
                                      If first_income = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Client"
                                          ' Text 360, y_pos, 55, 10, "Employer (or Business Name)"
                                          ' Text 420, y_pos, 30, 10, "Change"
                                          ' Text 455, y_pos, 35, 10, "Amount"
                                          Text 300, y_pos, 200, 10, "Client    &    Employer    &    Change    &    Amount"
                                          y_pos = y_pos + 10
                                          first_income = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_EARNED_ARRAY(earned_client, the_earned) & " - " & NEW_EARNED_ARRAY(earned_source, the_earned) & " - " & NEW_EARNED_ARRAY(earned_change_date, the_earned) & " - " & NEW_EARNED_ARRAY(earned_amount, the_earned)
                                      ' Text 300, y_pos, 55, 45, NEW_EARNED_ARRAY(earned_client, the_earned)
                                      ' Text 360, y_pos, 55, 15, NEW_EARNED_ARRAY(earned_source, the_earned)
                                      ' Text 420, y_pos, 30, 15, NEW_EARNED_ARRAY(earned_change_date, the_earned)
                                      ' Text 455, y_pos, 35, 15, NEW_EARNED_ARRAY(earned_amount, the_earned)
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_income = TRUE Then
                                  Text 300, y_pos, 200, 10, "No income entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 40, 10, "Q16. Notes:"
                              EditBox 325, y_pos, 185, 15, question_sixteen_notes

                              Text 10, 295, 50, 10, "JOBS Notes:"
                              EditBox 65, 290, 450, 15, ALL_INCOME_ARRAY(item_notes_const, ei_to_use)
                          End If
                      ElseIf panel_indicator = "UNEA" Then
                          If panel_array_to_use <> "" Then
                              ui_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)

                              CheckBox 280, 25, 145, 10, "Check Here if UNEA needs to be Verified.", ALL_INCOME_ARRAY(verif_checkbox_const, ui_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_INCOME_ARRAY(verif_time_const, ui_to_use)
                              Text 15, 20, 170, 10, "HH Member: " & ALL_INCOME_ARRAY(owner_name, ui_to_use)
                              Text 15, 35, 200, 10, "Unearned Income Type: " & ALL_INCOME_ARRAY(type_const, ui_to_use)
                              Text 15, 45, 140, 10, "Verification: " & ALL_INCOME_ARRAY(verif_const, ui_to_use)
                              Text 15, 55, 135, 10, "Claim Number: " & ALL_INCOME_ARRAY(claim_nbr_const, ui_to_use)
                              Text 15, 70, 110, 10, "Pay Frequency: " & ALL_INCOME_ARRAY(frequency_const, ui_to_use)
                              Text 15, 85, 250, 10, "Income Start: " & ALL_INCOME_ARRAY(start_date_const, ui_to_use) & "                  Income End: " & ALL_INCOME_ARRAY(end_date_const, ui_to_use)
                              Text 15, 100, 140, 10, "Retrospective:      Amount: $ " & ALL_INCOME_ARRAY(retro_income_amount, ui_to_use)
                              Text 15, 115, 140, 10, "Prospective:          Amount: $ " & ALL_INCOME_ARRAY(amount_const, ui_to_use)
                              GroupBox 15, 135, 220, 60, "SNAP PIC"
                              Text 25, 150, 100, 10, "Pay Frequency: " & ALL_INCOME_ARRAY(snap_pic_frequency, ui_to_use)
                              Text 35, 165, 175, 10, "Ave Income per Pay Date:   $ " & ALL_INCOME_ARRAY(snap_pic_income_per_pay, ui_to_use)
                              Text 25, 175, 175, 10, "Prospective Monthly Amount:   $ " & ALL_INCOME_ARRAY(snap_pic_monthly_income, ui_to_use)

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q7. Anyone have other income - " & ma_other_income
                              y_pos = y_pos + 10

                              first_unea = TRUE
                              For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
                                  If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
                                      If first_unea = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Name"
                                          ' Text 360, y_pos, 55, 10, "Type of Income"
                                          ' Text 420, y_pos, 35, 10, "Start Date"
                                          ' Text 460, y_pos, 35, 10, "Amount"
                                          Text 300, y_pos, 200, 10, "Name    &    Type of Income    &    Start Date    &    Amount"
                                          y_pos = y_pos + 10
                                          first_unea = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_UNEARNED_ARRAY(unearned_client, each_unea) & " - " & NEW_UNEARNED_ARRAY(unearned_source, each_unea) & " - " & NEW_UNEARNED_ARRAY(unearned_start_date, each_unea) & " - " & NEW_UNEARNED_ARRAY(unearned_amount, each_unea)
                                      ' Text 300, y_pos, 55, 45, NEW_UNEARNED_ARRAY(unearned_client, each_unea)
                                      ' Text 360, y_pos, 55, 45, NEW_UNEARNED_ARRAY(unearned_source, each_unea)   'unea_type
                                      ' Text 420, y_pos, 35, 15, NEW_UNEARNED_ARRAY(unearned_start_date, each_unea)    'unea_start_date
                                      ' Text 460, y_pos, 35, 15, NEW_UNEARNED_ARRAY(unearned_amount, each_unea)    'unea_amount
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_unea = TRUE Then
                                  Text 300, y_pos, 200, 10, "No UNEA entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 35, 10, "Q7 Notes:"
                              EditBox 325, y_pos, 185, 15, question_seven_notes
                              y_pos = y_pos + 49

                              Text 285, y_pos, 200, 10, "Q17. Has there been a change unearned income - " & quest_seventeen_form_answer
                              y_pos = y_pos + 10
                              first_unea = TRUE
                              For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
                                  If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
                                      If first_unea = TRUE Then
                                          ' Text 300, y_pos, 55, 10, "Client"
                                          ' Text 360, y_pos, 55, 10, "Type and Source"
                                          ' Text 420, y_pos, 35, 10, "Change Date"
                                          ' Text 460, y_pos, 35, 10, "Amount"
                                          Text 300, y_pos, 200, 10, "Client    &    Type and Source    &    Change Date    &    Amount"
                                          first_unea = FALSE
                                          y_pos = y_pos + 10
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_UNEARNED_ARRAY(unearned_client, the_unearned) & " - " & NEW_UNEARNED_ARRAY(unearned_source, the_unearned) & " - " & NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned) & " - " & NEW_UNEARNED_ARRAY(unearned_amount, the_unearned)
                                      ' Text 300, y_pos, 55, 45, NEW_UNEARNED_ARRAY(unearned_client, the_unearned)
                                      ' Text 360, y_pos, 55, 15, NEW_UNEARNED_ARRAY(unearned_source, the_unearned)
                                      ' Text 420, y_pos, 35, 15, NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned)
                                      ' Text 460, y_pos, 35, 15, NEW_UNEARNED_ARRAY(unearned_amount, the_unearned)
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_unea = TRUE Then
                                  Text 300, y_pos, 200, 15, "No UNEA entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 40, 10, "Q17. Notes:"
                              EditBox 325, y_pos, 185, 15, question_seventeen_notes
                              y_pos = y_pos + 40

                              Text 285, y_pos, 200, 10, "Q18. Has there been a change child support income - " & quest_eighteen_form_answer
                              y_pos = y_pos + 10
                              first_cses = TRUE
                              For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
                                  If NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) <> "Select..." Then
                                      If first_cses = TRUE Then
                                          ' Text 300, y_pos, 85, 10, "Name of person paying"
                                          ' Text 390, y_pos, 35, 10, "Amount"
                                          ' Text 435, y_pos, 65, 10, "Currently Paying?"
                                          Text 300, y_pos, 200, 10, "Name of person paying    &    Amount    &    Currently Paying"
                                          y_pos = y_pos + 10
                                          first_cses = FALSE
                                      End If
                                      Text 300, y_pos, 200, 10, NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) & " - " & NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs) & " - " & NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs)
                                      ' Text 300, y_pos, 85, 15, NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs)
                                      ' Text 390, y_pos, 35, 15, NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs)
                                      ' Text 435, y_pos, 65, 45, NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs)
                                      y_pos = y_pos + 10
                                  End If
                              Next
                              If first_cses = TRUE Then
                                  Text 300, y_pos, 200, 10, "No Child Support Income entered."
                                  y_pos = y_pos + 10
                              End If
                              Text 285, y_pos + 5, 40, 10, "Q18. Notes:"
                              EditBox 325, y_pos, 185, 15, question_eighteen_notes

                              Text 10, 295, 50, 10, "UNEA Notes:"
                              EditBox 65, 290, 450, 15, ALL_INCOME_ARRAY(item_notes_const, ui_to_use)
                          End If
                      ElseIf panel_indicator = "SHEL" Then
                          If panel_array_to_use <> "" Then
                              pers_to_use = ALL_PANELS_ARRAY(array_ref_const, panel_array_to_use)
                              CheckBox 280, 25, 145, 10, "Check Here if SHEL needs to be Verified.", ALL_CLIENTS_ARRAY(shel_verif_checkbox, pers_to_use)
                              Text 345, 40, 75, 10, "Time period to verify:"
                              EditBox 425, 35, 85, 15, ALL_CLIENTS_ARRAY(shel_verif_time, pers_to_use)

                              Text 15, 25, 200, 10, "Shelter Expense for: " & ALL_CLIENTS_ARRAY(memb_first_name, pers_to_use) & " " & ALL_CLIENTS_ARRAY(memb_last_name, pers_to_use)
                              Text 20, 40, 105, 10, "Shelter Expense Shared: " & ALL_CLIENTS_ARRAY(shel_shared_yn, pers_to_use)
                              Text 20, 50, 135, 10, "Shelter Expense HUD Subsidized: " & ALL_CLIENTS_ARRAY(shel_hud_sub_yn, pers_to_use)
                              Text 20, 65, 185, 10, "Landlord: " & ALL_CLIENTS_ARRAY(shel_paid_to, pers_to_use)
                              ' GroupBox 50, 75, 110, 150, "Retrospective"
                              Text 75, 80, 105, 10, "Amount                    Verification"
                              ' Text 100, 85, 40, 10, "Verification"
                              ' GroupBox 165, 75, 110, 150, "Prospective"
                              ' Text 170, 85, 105, 10, "Amount                    Verification"
                              ' Text 215, 85, 40, 10, "Verification"
                              If ALL_CLIENTS_ARRAY(shel_rent_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_rent_prosp_amt, pers_to_use) <> "" Then
                                  Text 25, 90, 250, 10, "Rent:   Retro -   $ " & ALL_CLIENTS_ARRAY(shel_rent_retro_amt, pers_to_use)         & "                    " & ALL_CLIENTS_ARRAY(shel_rent_retro_verif, pers_to_use)
                                  Text 48, 100, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_rent_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_rent_prosp_verif, pers_to_use)
                              Else
                                  Text 25, 90, 250, 10, "Rent:         None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_lot_rent_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_amt, pers_to_use) <> "" Then
                                  Text 15, 115, 250, 10, "Lot Rent:  Retro -   $" & ALL_CLIENTS_ARRAY(shel_lot_rent_retro_amt, pers_to_use)     & "                    " & ALL_CLIENTS_ARRAY(shel_lot_rent_retro_verif, pers_to_use)
                                  Text 48, 125, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_verif, pers_to_use)
                              Else
                                  Text 15, 115, 250, 10, "Lot Rent:             None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_mortgage_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_mortgage_prosp_amt, pers_to_use) <> "" Then
                                  Text 10, 140, 250, 10, "Mortgage:   Retro -   $" & ALL_CLIENTS_ARRAY(shel_mortgage_retro_amt, pers_to_use)     & "                    " & ALL_CLIENTS_ARRAY(shel_mortgage_retro_verif, pers_to_use)
                                  Text 48, 150, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_mortgage_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_mortgage_prosp_verif, pers_to_use)
                              Else
                                  Text 10, 140, 250, 10, "Mortgage:              None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_insurance_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_insurance_prosp_amt, pers_to_use) <> "" Then
                                  Text 10, 165, 250, 10, "Insurance:  Retro -   $" & ALL_CLIENTS_ARRAY(shel_insurance_retro_amt, pers_to_use)    & "                    " & ALL_CLIENTS_ARRAY(shel_insurance_retro_verif, pers_to_use)
                                  Text 48, 175, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_insurance_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_insurance_prosp_verif, pers_to_use)
                              Else
                                  Text 10, 165, 250, 10, "Insurance:             None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_tax_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_tax_prosp_amt, pers_to_use) <> "" Then
                                  Text 20, 190, 250, 10, "Taxes:   Retro -   $" & ALL_CLIENTS_ARRAY(shel_tax_retro_amt, pers_to_use)          & "                    " & ALL_CLIENTS_ARRAY(shel_tax_retro_verif, pers_to_use)
                                  Text 48, 200, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_tax_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_tax_prosp_verif, pers_to_use)
                              Else
                                  Text 20, 190, 250, 10, "Taxes:              None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_room_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_room_prosp_amt, pers_to_use) <> "" Then
                                  Text 20, 215, 250, 10, "Room:   Retro -   $" & ALL_CLIENTS_ARRAY(shel_room_retro_amt, pers_to_use)         & "                    " & ALL_CLIENTS_ARRAY(shel_room_retro_verif, pers_to_use)
                                  Text 48, 225, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_room_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_room_prosp_verif, pers_to_use)
                              Else
                                  Text 20, 215, 250, 10, "Room:              None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_garage_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_garage_prosp_amt, pers_to_use) <> "" Then
                                  Text 15, 240, 250, 10, "Garage:    Retro -   $" & ALL_CLIENTS_ARRAY(shel_garage_retro_amt, pers_to_use)       & "                    " & ALL_CLIENTS_ARRAY(shel_garage_retro_verif, pers_to_use)
                                  Text 48, 250, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_garage_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_garage_prosp_verif, pers_to_use)
                              Else
                                  Text 15, 240, 250, 10, "Garage:               None"
                              End If

                              If ALL_CLIENTS_ARRAY(shel_subsidy_retro_amt, pers_to_use) <> "" OR ALL_CLIENTS_ARRAY(shel_subsidy_prosp_amt, pers_to_use) <> "" Then
                                  Text 15, 265, 250, 10, "Subsidy:    Retro -   $" & ALL_CLIENTS_ARRAY(shel_subsidy_retro_amt, pers_to_use)      & "                    " & ALL_CLIENTS_ARRAY(shel_subsidy_retro_verif, pers_to_use)
                                  Text 48, 275, 220, 10, "Prosp -   $ " & ALL_CLIENTS_ARRAY(shel_subsidy_prosp_amt, pers_to_use)        & "                    " & ALL_CLIENTS_ARRAY(shel_subsidy_prosp_verif, pers_to_use)
                              Else
                                  Text 15, 265, 250, 10, "Subsidy:               None"
                              End If


                              ' Text 55, 100, 105, 10, ALL_CLIENTS_ARRAY(shel_rent_retro_amt, pers_to_use)         & "          " & ALL_CLIENTS_ARRAY(shel_rent_retro_verif, pers_to_use)
                              ' Text 55, 115, 105, 10, ALL_CLIENTS_ARRAY(shel_lot_rent_retro_amt, pers_to_use)     & "          " & ALL_CLIENTS_ARRAY(shel_lot_rent_retro_verif, pers_to_use)
                              ' Text 55, 130, 105, 10, ALL_CLIENTS_ARRAY(shel_mortgage_retro_amt, pers_to_use)     & "          " & ALL_CLIENTS_ARRAY(shel_mortgage_retro_verif, pers_to_use)
                              ' Text 55, 145, 105, 10, ALL_CLIENTS_ARRAY(shel_insurance_retro_amt, pers_to_use)    & "          " & ALL_CLIENTS_ARRAY(shel_insurance_retro_verif, pers_to_use)
                              ' Text 55, 160, 105, 10, ALL_CLIENTS_ARRAY(shel_tax_retro_amt, pers_to_use)          & "          " & ALL_CLIENTS_ARRAY(shel_tax_retro_verif, pers_to_use)
                              ' Text 55, 175, 105, 10, ALL_CLIENTS_ARRAY(shel_room_retro_amt, pers_to_use)         & "          " & ALL_CLIENTS_ARRAY(shel_room_retro_verif, pers_to_use)
                              ' Text 55, 190, 105, 10, ALL_CLIENTS_ARRAY(shel_garage_retro_amt, pers_to_use)       & "          " & ALL_CLIENTS_ARRAY(shel_garage_retro_verif, pers_to_use)
                              ' Text 55, 210, 105, 10, ALL_CLIENTS_ARRAY(shel_subsidy_retro_amt, pers_to_use)      & "          " & ALL_CLIENTS_ARRAY(shel_subsidy_retro_verif, pers_to_use)

                              ' Text 100, 100, 60, 10, rent_retro_verif
                              ' Text 100, 115, 60, 10, lot_rent_retro_verif
                              ' Text 100, 130, 60, 10, mortgage_retro_verif
                              ' Text 100, 145, 60, 10, insurance_retro_verif
                              ' Text 100, 160, 60, 10, tax_retro_verif
                              ' Text 100, 175, 60, 10, room_retro_verif
                              ' Text 100, 190, 60, 10, garage_retro_verif
                              ' Text 100, 210, 60, 10, subsidy_retro_verif


                              ' Text 170, 100, 105, 10, ALL_CLIENTS_ARRAY(shel_rent_prosp_amt, pers_to_use)        & "          " & ALL_CLIENTS_ARRAY(shel_rent_prosp_verif, pers_to_use)
                              ' Text 170, 115, 105, 10, ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_amt, pers_to_use)    & "          " & ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_verif, pers_to_use)
                              ' Text 170, 130, 105, 10, ALL_CLIENTS_ARRAY(shel_mortgage_prosp_amt, pers_to_use)    & "          " & ALL_CLIENTS_ARRAY(shel_mortgage_prosp_verif, pers_to_use)
                              ' Text 170, 145, 105, 10, ALL_CLIENTS_ARRAY(shel_insurance_prosp_amt, pers_to_use)   & "          " & ALL_CLIENTS_ARRAY(shel_insurance_prosp_verif, pers_to_use)
                              ' Text 170, 160, 105, 10, ALL_CLIENTS_ARRAY(shel_tax_prosp_amt, pers_to_use)         & "          " & ALL_CLIENTS_ARRAY(shel_tax_prosp_verif, pers_to_use)
                              ' Text 170, 175, 105, 10, ALL_CLIENTS_ARRAY(shel_room_prosp_amt, pers_to_use)        & "          " & ALL_CLIENTS_ARRAY(shel_room_prosp_verif, pers_to_use)
                              ' Text 170, 190, 105, 10, ALL_CLIENTS_ARRAY(shel_garage_prosp_amt, pers_to_use)      & "          " & ALL_CLIENTS_ARRAY(shel_garage_prosp_verif, pers_to_use)
                              ' Text 170, 210, 105, 10, ALL_CLIENTS_ARRAY(shel_subsidy_prosp_amt, pers_to_use)     & "          " & ALL_CLIENTS_ARRAY(shel_subsidy_prosp_verif, pers_to_use)

                              ' Text 215, 100, 60, 10, rent_prosp_verif
                              ' Text 215, 115, 60, 10, lot_rent_prosp_verif
                              ' Text 215, 130, 60, 10, mortgage_prosp_verif
                              ' Text 215, 145, 60, 10, insurance_prosp_verif
                              ' Text 215, 160, 60, 10, tax_prosp_verif
                              ' Text 215, 175, 60, 10, room_prosp_verif
                              ' Text 215, 190, 60, 10, garage_prosp_verif
                              ' Text 215, 210, 60, 10, subsidy_prosp_verif
                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q15. Anyone have other income - " & quest_fifteen_form_answer
                              y_pos = y_pos + 10
                              If new_rent_or_mortgage_amount <> "" Then Text 300, y_pos, 150, 10, "New Rent or Mortgage Amount: " & new_rent_or_mortgage_amount
                              If new_rent_or_mortgage_amount = "" Then Text 300, y_pos, 150, 10, "New Rent or Mortgage Amount: BLANK"
                              y_pos = y_pos + 10
                              util_checked = ""
                              If heat_ac_checkbox = checked Then util_checked = util_checked & " & Heat/AC"
                              If electricity_checkbox = checked Then util_checked = util_checked & " & Electric"
                              If telephone_checkbox = checked Then util_checked = util_checked & " & Phone"
                              If util_checked = "" Then
                                  util_checked = "NO UTILITIES"
                              Else
                                  util_checked = right(util_checked, len(util_checked) - 3)
                              End If
                              Text 300, y_pos, 150, 10, "Client indicated they pay: " & util_checked
                              y_pos = y_pos + 10
                              If shel_proof_provided <> "Select One..." Then Text 300, y_pos, 150, 10, "Proof attached: " & shel_proof_provided
                              y_pos = y_pos + 10
                              Text 285, y_pos + 5, 40, 10, "Q15 Notes:"
                              EditBox 325, y_pos, 185, 15, question_fifteen_notes

                              Text 10, 295, 50, 10, "SHEL Notes:"
                              EditBox 65, 290, 450, 15, ALL_CLIENTS_ARRAY(shel_notes, pers_to_use)
                          End If
                      ElseIf panel_indicator = "HEST" Then
                          If panel_array_to_use <> "" Then
                              CheckBox 280, 25, 145, 10, "Check Here if more information about Utility Expense is needed.", HEST_verif_checkbox

                              GroupBox 10, 20, 230, 40, "THIS CASE COUNTED UTILITY EXPENSE"
                              Text 25, 40, 190, 10, "Utilities Expense (standard allowance) $ " & HEST_total_expense
                              Text 15, 70, 210, 10, "Persons Paying: " & HEST_persons_paying
                              Text 15, 85, 100, 10, "FS Choice Date: " & HEST_fs_choice_date
                              Text 15, 100, 165, 10, "Actual Expense in Initial Month: $ " & HEST_initial_month_actual_expense
                              Text 20, 130, 30, 10, "Heat/Air:"
                              Text 20, 145, 30, 10, "Electric:"
                              Text 25, 160, 25, 10, "Phone:"
                              Text 60, 115, 50, 10, "Retrospective"
                              Text 135, 115, 50, 10, "Prospective"
                              If HEST_retro_heat_air <> "" Then Text 60, 130, 50, 10, HEST_retro_heat_air & " - $ " & HEST_retro_heat_air_amount
                              If HEST_retro_electric <> "" Then Text 60, 145, 50, 10, HEST_retro_electric & " - $ " & HEST_retro_electric_amount
                              If HEST_retro_phone <> "" Then Text 60, 160, 50, 10, HEST_retro_phone & " - $ " & HEST_retro_phone_amount
                              If HEST_prosp_heat_air <> "" Then Text 135, 130, 50, 10, HEST_prosp_heat_air & " - $ " & HEST_prosp_heat_air_amount
                              If HEST_prosp_electric <> "" Then Text 135, 145, 50, 10, HEST_prosp_electric & " - $ " & HEST_prosp_electric_amount
                              If HEST_prosp_phone <> "" Then Text 135, 160, 50, 10, HEST_prosp_phone & " - $ " & HEST_prosp_phone_amount

                              y_pos = 55
                              Text 285, y_pos, 200, 10, "Q15. Anyone have other income - " & quest_fifteen_form_answer
                              y_pos = y_pos + 10
                              If new_rent_or_mortgage_amount <> "" Then Text 300, y_pos, 150, 10, "New Rent or Mortgage Amount: " & new_rent_or_mortgage_amount
                              If new_rent_or_mortgage_amount = "" Then Text 300, y_pos, 150, 10, "New Rent or Mortgage Amount: BLANK"
                              y_pos = y_pos + 10
                              util_checked = ""
                              If heat_ac_checkbox = checked Then util_checked = util_checked & " & Heat/AC"
                              If electricity_checkbox = checked Then util_checked = util_checked & " & Electric"
                              If telephone_checkbox = checked Then util_checked = util_checked & " & Phone"
                              If util_checked = "" Then
                                  util_checked = "NO UTILITIES"
                              Else
                                  util_checked = right(util_checked, len(util_checked) - 3)
                              End If
                              Text 300, y_pos, 150, 10, "Client indicated they pay: " & util_checked
                              y_pos = y_pos + 10
                              If shel_proof_provided <> "Select One..." Then Text 300, y_pos, 150, 10, "Proof attached: " & shel_proof_provided
                              y_pos = y_pos + 10
                              Text 285, y_pos + 5, 40, 10, "Q15 Notes:"
                              EditBox 325, y_pos, 185, 15, question_fifteen_notes

                              Text 10, 295, 50, 10, "HEST Notes:"
                              EditBox 65, 290, 450, 15, hest_notes

                          End If
                      ElseIf panel_indicator = "NOTES" Then
                          notes_reviewed = TRUE

                          csr_status_grp_len = 10
                          If snap_sr_yn = "Yes" Then csr_status_grp_len = csr_status_grp_len + 70
                          If hc_sr_yn = "Yes" Then csr_status_grp_len = csr_status_grp_len + 70
                          If grh_sr_yn = "Yes" Then csr_status_grp_len = csr_status_grp_len + 70
                          GroupBox 10, 15, 500, csr_status_grp_len, "CSR Status"
                          y_pos = 30
                          SNAP_verifs_needed = "NOT Needed"
                          HC_verifs_needed = "NOT Needed"
                          GRH_verifs_needed = "NOT Needed"
                          If trim(verifs_needed) <> "" Then
                              If verifs_needed_for_SNAP_checkbox = checked Then SNAP_verifs_needed = "Needed"
                              If verifs_needed_for_HC_checkbox = checked Then HC_verifs_needed = "Needed"
                              If verifs_needed_for_GRH_checkbox = checked Then GRH_verifs_needed = "Needed"
                          End If

                          If snap_sr_yn = "Yes" Then
                              If snap_sr_status <> "Select One..." Then
                                  If snap_questions_complete = TRUE AND SNAP_verifs_needed = "NOT Needed" Then snap_sr_status = "U - Complete and Updt Req"
                                  If snap_questions_complete = FALSE Then snap_sr_status = "I - Incomplete"
                                  If SNAP_verifs_needed = "Needed" Then snap_sr_status = "I - Incomplete"
                              End If
                              Text 15, y_pos, 175, 10, "For the " & snap_sr_mo & "/" & snap_sr_yr & " SNAP SR, the status appears: "
                              DropListBox 155, y_pos-5, 90, 45, "Select One..."+chr(9)+"I - Incomplete"+chr(9)+"U - Complete and Updt Req"+chr(9)+"N - Not Rcvd"+chr(9)+"A - Approved"+chr(9)+"O - Override Autoclose"+chr(9)+"T - Terminated"+chr(9)+"D - Denied", snap_sr_status
                              If snap_questions_complete = TRUE Then Text 15, y_pos + 15, 300, 10, "This status is based upon the following:    - Form is COMPLETE"
                              If snap_questions_complete = FALSE Then Text 15, y_pos + 15, 300, 10, "This status is based upon the following:    - Form is INCOMPLETE"
                              Text 145, y_pos + 25, 100, 10, " - Verifications: " & SNAP_verifs_needed
                              Text 15, y_pos + 45, 60, 10, "SNAP SR Notes:"
                              EditBox 80, y_pos + 40, 400, 15, snap_sr_notes
                              y_pos = y_pos + 70
                          End If

                          If hc_sr_yn = "Yes" Then
                              If hc_sr_status <> "Select One..." Then
                                  If ma_questions_complete = TRUE AND HC_verifs_needed = "NOT Needed" Then hc_sr_status = "U - Complete and Updt Req"
                                  If ma_questions_complete = FALSE Then hc_sr_status = "I - Incomplete"
                                  If HC_verifs_needed = "Needed" Then hc_sr_status = "I - Incomplete"
                              End If
                              Text 15, y_pos, 175, 10, "For the " & hc_sr_mo & "/" & hc_sr_yr & " HC SR, the status appears: "
                              DropListBox 155, y_pos-5, 90, 45, "Select One..."+chr(9)+"I - Incomplete"+chr(9)+"U - Complete and Updt Req"+chr(9)+"N - Not Rcvd"+chr(9)+"A - Approved"+chr(9)+"O - Override Autoclose"+chr(9)+"T - Terminated"+chr(9)+"D - Denied", hc_sr_status
                              If ma_questions_complete = TRUE THen Text 15, y_pos + 15, 300, 10, "This status is based upon the following:    - Form is COMPLETE"
                              If ma_questions_complete = FALSE Then Text 15, y_pos + 15, 300, 10, "This status is based upon the following:    - Form is INCOMPLETE"
                              Text 145, y_pos + 25, 100, 10, " - Verifications: " & HC_verifs_needed
                              Text 15, y_pos + 45, 60, 10, "HC SR Notes:"
                              EditBox 80, y_pos + 40, 400, 15, hc_sr_notes
                              y_pos = y_pos + 70

                          End If

                          If grh_sr_yn = "Yes" Then
                              If grh_sr_status <> "Select One..." Then
                                  If grh_questions_complete = TRUE AND GRH_verifs_needed = "NOT Needed" Then grh_sr_status = "U - Complete and Updt Req"
                                  If grh_questions_complete = FALSE Then grh_sr_status = "I - Incomplete"
                                  If GRH_verifs_needed = "Needed" Then grh_sr_status = "I - Incomplete"
                              End If
                              Text 15, y_pos, 175, 10, "For the " & grh_sr_mo & "/" & grh_sr_yr & " GRH SR, the status appears: "
                              DropListBox 155, y_pos-5, 90, 45, "Select One..."+chr(9)+"I - Incomplete"+chr(9)+"U - Complete and Updt Req"+chr(9)+"N - Not Rcvd"+chr(9)+"A - Approved"+chr(9)+"O - Override Autoclose"+chr(9)+"T - Terminated"+chr(9)+"D - Denied", grh_sr_status
                              If grh_questions_complete = TRUE Then Text 15, y_pos + 15, 300, 10, "This status is based upon the following:    - Form is COMPLETE"
                              If grh_questions_complete = FALSE Then Text 15, y_pos + 15, 300, 10, "This status is based upon the following:    - Form is INCOMPLETE"
                              Text 145, y_pos + 25, 100, 10, " - Verifications: " & GRH_verifs_needed
                              Text 15, y_pos + 45, 60, 10, "GRH SR Notes:"
                              EditBox 80, y_pos + 40, 400, 15, grh_sr_notes
                              y_pos = y_pos + 70
                          End If

                          GroupBox 10, y_pos, 500, 70, "Check any Boxes that apply"
                          CheckBox 20, y_pos + 20, 200, 10, "CSR Cash Supplement used as HRF", HRF_checkbox
                          CheckBox 20, y_pos + 30, 200, 10, "Checked eDRS", eDRS_sent_checkbox
                          CheckBox 20, y_pos + 40, 200, 10, "Forms sent to AREP", Sent_arep_checkbox
                          CheckBox 20, y_pos + 50, 200, 10, "Emailed MADE through SIR", MADE_checkbox

                          ' Text 15, 25, 200, 10, ""
                          ' GroupBox 10, 15, 500, 70, "General CSR Questions"

                          ' Text 15, 25, 200, 10, "Q1. Name listed on CSR form - " & client_on_csr_form
                          ' Text 220, 25, 35, 10, "Q1 Notes:"
                          ' EditBox 260, 20, 230, 15, question_one_notes
                          ' Text 15, 35, 400, 10, "      Residence address on CSR form - " & residence_address_match_yn & "      Mailing address on CSR form - " & mailing_address_match_yn
                          ' ' Text 15, 45, 200, 10, "      Mailing address on CSR form - " & mailing_address_match_yn
                          ' Text 15, 45, 200, 10, "      Homeless status on CSR form - " & homeless_status
                          '
                          ' Text 15, 60, 200, 10, "Q2. Has anyone moved in or out - " & quest_two_move_in_out
                          ' Text 220, 60, 35, 10, "Q2 Notes:"
                          ' EditBox 260, 55, 230, 15, question_two_notes
                          ' clients_moved_in = ""
                          ' clients_moved_out = ""
                          ' For known_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
                          '     If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then clients_moved_out = clients_moved_out & "Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & ", " & vbCr
                          '     If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then clients_moved_in = clients_moved_in & "Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & ", " & vbCr
                          ' Next
                          ' If clients_moved_out = "" Then clients_moved_out = "None"
                          ' If clients_moved_in  ="" Then clients_moved_in = "None"
                          ' Text 15, 70, 400, 10, "      Clients moved out: " & clients_moved_out & "      Clients moved in: " & clients_moved_in
                          ' Text 15, 90, 200, 10, "      Clients moved in: " & clients_moved_in

                          ' GroupBox 10, 80, 500, 120, "MA CSR Questions"


                          ' GroupBox 10, 195, 500, 80, "SNAP CSR Questions"

                          Text 10, 275, 50, 10, "Action Taken:"
                          EditBox 65, 270, 450, 15, actions_taken
                          Text 10, 295, 50, 10, "Other Notes:"
                          EditBox 65, 290, 450, 15, other_notes
                      ElseIf panel_indicator = "VERIFS" Then
                          verif_reviewed = TRUE
                          ' verifs_needed = "This is a long thing with lots of words because I need to see how the words will fit in the dialog and how it will look when the VERIFICATIONS button is pressed in the CSR final dialog because that is important for the functioning of this script and I am going to just keep typing. My eyes are closed right now because it doesn't even really matter what I am typing or how well it is written. I say that but then i jest HAD to check if I had spelled matter the right way. OK, I think this is enough for now."
                          ' Text 10, 20, 120, 10, "Date Verification Request Form Sent:"
                          Text 10, 20, 500, 40, "Date Verification Request Form Sent:                                                                                                                                                                                                      Check the boxes for any verification you want to add to the CASE/NOTE.                                                                                                                                Note: After you press 'Fill' or any of the other panel buttons the information from the boxes will fill in the Verification Field and the boxes will be 'unchecked'."
                          ' Text 10, 30, 500, 20, "Check the boxes for any verification you want to add to the CASE/NOTE.                                                                                                                                Note: After you press 'Fill' or any of the other panel buttons the information from the boxes will fill in the Verification Field and the boxes will be 'unchecked'."
                          EditBox 135, 15, 50, 15, verif_req_form_sent_date

                          GroupBox 5, 50, 515, 70, "Personal and Household Information"

                          ' CheckBox 10, 50, 75, 10, "Verification of ID for ", id_verif_checkbox
                          ' ComboBox 90, 45, 150, 45, verification_memb_list, id_verif_memb
                          ' CheckBox 300, 50, 100, 10, "Social Security Number for ", ssn_checkbox
                          ' ComboBox 405, 45, 150, 45, verification_memb_list, ssn_verif_memb
                          '
                          ' CheckBox 10, 70, 70, 10, "US Citizenship for ", us_cit_status_checkbox
                          ' ComboBox 85, 65, 150, 45, verification_memb_list, us_cit_verif_memb
                          ' CheckBox 300, 70, 85, 10, "Immigration Status for", imig_status_checkbox
                          ' ComboBox 390, 65, 150, 45, verification_memb_list, imig_verif_memb
                          '
                          ' CheckBox 10, 90, 90, 10, "Proof of relationship for ", relationship_checkbox
                          ' ComboBox 105, 85, 150, 45, verification_memb_list, relationship_one_verif_memb
                          ' Text 260, 90, 90, 10, "and"
                          ' ComboBox 280, 85, 150, 45, verification_memb_list, relationship_two_verif_memb

                          CheckBox 10, 65, 85, 10, "Student Information for ", student_info_checkbox
                          ComboBox 100, 60, 150, 45, verification_memb_list, student_verif_memb
                          Text 255, 65, 10, 10, "at"
                          EditBox 270, 60, 150, 15, student_verif_source

                          CheckBox 10, 85, 85, 10, "Proof of Pregnancy for", preg_checkbox
                          ComboBox 100, 80, 150, 45, verification_memb_list, preg_verif_memb

                          CheckBox 10, 105, 115, 10, "Illness/Incapacity/Disability for", illness_disability_checkbox
                          ComboBox 130, 100, 150, 45, verification_memb_list, disa_verif_memb
                          Text 285, 105, 30, 10, "verifying:"
                          EditBox 320, 100, 150, 15, disa_verif_type

                          GroupBox 5, 120, 515, 50, "Income Information"

                          CheckBox 10, 135, 45, 10, "Income for ", income_checkbox
                          ComboBox 60, 130, 130, 45, verification_memb_list, income_verif_memb
                          Text 195, 135, 200, 10, "from                                                                       for"
                          ComboBox 215, 130, 130, 45, income_source_list, income_verif_source
                          ' Text 370, 135, 10, 10, "for"
                          EditBox 365, 130, 140, 15, income_verif_time

                          CheckBox 10, 155, 85, 10, "Employment Status for ", employment_status_checkbox
                          ComboBox 100, 150, 150, 45, verification_memb_list, emp_status_verif_memb
                          Text 255, 155, 10, 10, "at"
                          ComboBox 270, 150, 150, 45, employment_source_list, emp_status_verif_source

                          GroupBox 5, 170, 515, 50, "Expense Information"

                          CheckBox 10, 185, 105, 10, "Educational Funds/Costs for", educational_funds_cost_checkbox
                          ComboBox 120, 180, 150, 45, verification_memb_list, stin_verif_memb

                          CheckBox 10, 205, 65, 10, "Shelter Costs for ", shelter_checkbox
                          ComboBox 80, 200, 150, 45, verification_memb_list, shelter_verif_memb
                          checkBox 240, 205, 175, 10, "Check here if this verif is NOT MANDATORY", shelter_not_mandatory_checkbox

                          GroupBox 5, 220, 515, 30, "Asset Information"

                          CheckBox 10, 235, 70, 10, "Bank Account for", bank_account_checkbox
                          ComboBox 80, 230, 135, 45, verification_memb_list, bank_verif_memb
                          Text 220, 235, 200, 10, "account type                                                          for"
                          ComboBox 270, 230, 100, 45, "Select or Type"+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Cert of Deposit (CD)"+chr(9)+"Stock"+chr(9)+"Money Market", bank_verif_type
                          ' Text 405, 235, 10, 10, "for"
                          EditBox 390, 230, 120, 15, bank_verif_time

                          Text 10, 260, 20, 10, "Other:"
                          EditBox 30, 255, 480, 15, other_verifs
                          Checkbox 300, 275, 200, 10, "Check here to have verifs numbered in the CASE/NOTE.", number_verifs_checkbox
                          ' Checkbox 270, 275, 200, 10, "Check here if there are verifs that have been postponed.", verifs_postponed_checkbox

                          ButtonGroup ButtonPressed
                            PushButton 465, 15, 50, 15, "FILL", fill_button
                          Text 10, 275, 60, 10, "Verifs Requested:"
                          Text 10, 285, 505, 45, verifs_needed
                          ' Text 10, 30, 235, 10, "Check the boxes for any verification you want to add to the CASE/NOTE."
                          Text 10, 340, 230, 10, "Check the boxes for the programs these verifications are needed for:"
                          CheckBox 250, 340, 45, 10, "SNAP", verifs_needed_for_SNAP_checkbox
                          CheckBox 300, 340, 45, 10, "HC", verifs_needed_for_HC_checkbox
                          CheckBox 350, 340, 45, 10, "SNAP", verifs_needed_for_GRH_checkbox

                          CheckBox 10, 355, 300, 10, "Check here to have the verification information exported to a Word Document", export_verifs_info_to_work_checkbox
                      End If
                      If panel_indicator <> "VERIFS" Then
                          Text 10, 315, 50, 10, "Verifs Needed:"
                          EditBox 65, 310, 450, 15, verifs_needed
                      End If
                    EndDialog

                    dialog dialog1
                    ' MsgBox ButtonPressed & vbNewLine & "Next - " & next_btn & vbNewLine & "OK - " & OK

                    cancel_confirmation
                    If ButtonPressed = -1 Then ButtonPressed = next_btn
                    ' MsgBox ButtonPressed

                    err_msg = "LOOP" & err_msg

                    If panel_indicator = "VERIFS" Then
                        ' id_verif_checkbox = unchecked
                        ' us_cit_status_checkbox = unchecked
                        ' imig_status_checkbox = unchecked
                        ' ssn_checkbox = unchecked
                        ' relationship_checkbox = unchecked
                        ' income_checkbox = unchecked
                        ' employment_status_checkbox = unchecked
                        ' student_info_checkbox = unchecked
                        ' educational_funds_cost_checkbox = unchecked
                        ' shelter_checkbox = unchecked
                        ' bank_account_checkbox = unchecked
                        ' preg_checkbox = unchecked
                        ' illness_disability_checkbox = unchecked
                        verif_err_msg = ""

                        If income_checkbox = checked Then
                            If income_verif_memb = "Select or Type Member" OR trim(income_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose income needs to be verified."
                            If trim(income_verif_source) = "" OR trim(income_verif_source) = "Select or Type Source" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of income to be verified."
                            If trim(income_verif_time) = "[Enter Time Frame]" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the time frame of the income verification needed."
                        End If
                        If employment_status_checkbox = checked Then
                            If trim(emp_status_verif_source) = "" OR trim(emp_status_verif_source) = "Select or Type Source" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of the employment that needs status verified."
                            If emp_status_verif_memb = "Select or Type Member" OR trim(emp_status_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose employment status needs to be verified."
                        End If
                        If student_info_checkbox = checked Then
                            If trim(student_verif_source) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of school information to be verified"
                            If student_verif_memb = "Select or Type Member" OR trim(student_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member for which we need school verification."
                        End If
                        If educational_funds_cost_checkbox = checked AND (stin_verif_memb = "Select or Type Member" OR trim(stin_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member with educational funds and costs we need verified."
                        If shelter_checkbox = checked AND (shelter_verif_memb = "Select or Type Member" OR trim(shelter_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose shelter expense we need verified."
                        If bank_account_checkbox = checked Then
                            If trim(bank_verif_type) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the type of bank account to verify."
                            If bank_verif_memb = "Select or Type Member" OR trim(bank_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose bank account we need verified."
                            If trim(bank_verif_time) = "[Enter Time Frame]" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the time frame of the bank account verification needed."
                        End If
                        If preg_checkbox = checked AND (preg_verif_memb = "Select or Type Member" OR trim(preg_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose pregnancy needs to be verified."
                        If illness_disability_checkbox = checked Then
                            If trim(disa_verif_type) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the type (or details) of the illness/incapacity/disability that need to be verified."
                            If disa_verif_memb = "Select or Type Member" OR trim(disa_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose illness/incapacity/disability needs to be verified."
                        End If

                        If verifs_needed <> "" Then
                            If right(verifs_needed, 1) = ";" Then
                                verifs_needed = verifs_needed & " "
                            ElseIf right(verifs_needed, 2) <> "; " Then
                                verifs_needed = verifs_needed & "; "
                            End If
                        End If
                        If verif_err_msg = "" Then
                            If income_checkbox = checked Then
                                If IsNumeric(left(income_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Income for Memb " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                                Else
                                    verifs_needed = verifs_needed & "Income for " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                                End If
                                income_checkbox = unchecked
                                income_verif_source = ""
                                income_verif_memb = ""
                                income_verif_time = ""
                            End If
                            If employment_status_checkbox = checked Then
                                If IsNumeric(left(emp_status_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Employment Status for Memb " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                                Else
                                    verifs_needed = verifs_needed & "Employment Status for " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                                End If
                                employment_status_checkbox = unchecked
                                emp_status_verif_memb = ""
                                emp_status_verif_source = ""
                            End If
                            If student_info_checkbox = checked Then
                                If IsNumeric(left(student_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Student information for Memb " & student_verif_memb & " at " & student_verif_source & ".; "
                                Else
                                    verifs_needed = verifs_needed & "Student information for " & student_verif_memb & " at " & student_verif_source & ".; "
                                End If
                                student_info_checkbox = unchecked
                                student_verif_memb = ""
                                student_verif_source = ""
                            End If
                            If educational_funds_cost_checkbox = checked Then
                                If IsNumeric(left(stin_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Educational funds and costs for Memb " & stin_verif_memb & ".; "
                                Else
                                    verifs_needed = verifs_needed & "Educational funds and costs for " & stin_verif_memb & ".; "
                                End If
                                educational_funds_cost_checkbox = unchecked
                                stin_verif_memb = ""
                            End If
                            If shelter_checkbox = checked Then
                                If IsNumeric(left(shelter_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Shelter costs for Memb " & shelter_verif_memb & ". "
                                Else
                                    verifs_needed = verifs_needed & "Shelter costs for " & shelter_verif_memb & ". "
                                End If
                                If shelter_not_mandatory_checkbox = checked Then verifs_needed = verifs_needed & " THIS VERIFICATION IS NOT MANDATORY."
                                verifs_needed = verifs_needed & "; "
                                shelter_checkbox = unchecked
                                shelter_verif_memb = ""
                            End If
                            If bank_account_checkbox = checked Then
                                If IsNumeric(left(bank_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & bank_verif_type & " account for Memb " & bank_verif_memb & " for " & bank_verif_time & ".; "
                                Else
                                    verifs_needed = verifs_needed & bank_verif_type & " account for " & bank_verif_memb & " for " & bank_verif_time & ".; "
                                End If
                                bank_account_checkbox = unchecked
                                bank_verif_type = ""
                                bank_verif_memb = ""
                                bank_verif_time = ""
                            End If
                            If preg_checkbox = checked Then
                                If IsNumeric(left(preg_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Pregnancy for Memb " & preg_verif_memb & ".; "
                                Else
                                    verifs_needed = verifs_needed & "Pregnancy for " & preg_verif_memb & ".; "
                                End If
                                preg_checkbox = unchecked
                                preg_verif_memb = ""
                            End If
                            If illness_disability_checkbox = checked Then
                                If IsNumeric(left(disa_verif_memb, 2)) = TRUE Then
                                    verifs_needed = verifs_needed & "Ill/Incap or Disability for Memb " & disa_verif_memb & " of " & disa_verif_type & ",; "
                                Else
                                    verifs_needed = verifs_needed & "Ill/Incap or Disability for " & disa_verif_memb & " of " & disa_verif_type & ",; "
                                End If
                                illness_disability_checkbox = unchecked
                                disa_verif_memb = ""
                                disa_verif_type = ""
                            End If
                            other_verifs = trim(other_verifs)
                            If other_verifs <> "" Then verifs_needed = verifs_needed & other_verifs & "; "
                            other_verifs = ""
                        Else
                            MsgBox "Additional detail about verifications to note is needed:" & vbNewLine & verif_err_msg
                            ButtonPressed = verifs_tab_btn
                        End If
                    End If

                    find_indicator = TRUE
                    If ButtonPressed = previous_btn Then
                        If panel_array_to_use = "" Then
                            If panel_indicator = "ADDR" Then
                                ButtonPressed = address_tab_btn
                            ElseIf panel_indicator = "MEMB" Then
                                ButtonPressed = address_tab_btn
                            ElseIf panel_indicator = "WREG" Then
                                ButtonPressed = hh_comp_tab_btn
                            ElseIf panel_indicator = "FACI" Then
                                ButtonPressed = wreg_btn
                            ElseIf panel_indicator = "REVW" Then
                                ButtonPressed = faci_btn
                            ElseIf panel_indicator = "ACCT" Then
                                ButtonPressed = revw_btn
                            ElseIf panel_indicator = "SECU" Then
                                ButtonPressed = acct_tab_btn
                            ElseIf panel_indicator = "CARS" Then
                                ButtonPressed = secu_tab_btn
                            ElseIf panel_indicator = "REST" Then
                                ButtonPressed = cars_tab_btn
                            ElseIf panel_indicator = "UNEA" Then
                                ButtonPressed = rest_tab_btn
                            ElseIf panel_indicator = "JOBS" Then
                                ButtonPressed =  unea_tab_btn
                            ElseIf panel_indicator = "BUSI" Then
                                ButtonPressed = jobs_tab_btn
                            ElseIf panel_indicator = "SHEL" Then
                                ButtonPressed = busi_tabs_btn
                            ElseIf panel_indicator = "HEST" Then
                                ButtonPressed = shel_tab_btn
                            ElseIf panel_indicator = "VERIFS" Then
                                ButtonPressed = hst_tab_btn
                            ElseIf panel_indicator = "NOTES" Then
                                ButtonPressed = verifs_tab_btn
                            End If
                        ElseIf panel_array_to_use = 0 Then
                            ButtonPressed = address_tab_btn
                        Else
                            panel_array_to_use = panel_array_to_use - 1
                            Do
                                If ALL_PANELS_ARRAY(show_this_panel, panel_array_to_use) = FALSE Then panel_array_to_use = panel_array_to_use - 1
                            Loop until ALL_PANELS_ARRAY(show_this_panel, panel_array_to_use) = TRUE
                            If panel_indicator = ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use) Then find_indicator = FALSE
                            panel_indicator = ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use)
                            second_indicator = ""
                            if panel_indicator = "ACCT" Then second_indicator = "CASH"
                        End If
                    ElseIf ButtonPressed = next_btn Then
                        If panel_array_to_use = "" Then
                            If panel_indicator = "ADDR" Then
                                ButtonPressed = hh_comp_tab_btn
                            ElseIf panel_indicator = "MEMB" Then
                                ButtonPressed = wreg_btn
                            ElseIf panel_indicator = "WREG" Then
                                ButtonPressed = faci_btn
                            ElseIf panel_indicator = "FACI" Then
                                ButtonPressed = revw_btn
                            ElseIf panel_indicator = "REVW" Then
                                ButtonPressed = acct_tab_btn
                            ElseIf panel_indicator = "ACCT" Then
                                ButtonPressed =  secu_tab_btn
                            ElseIf panel_indicator = "SECU" Then
                                ButtonPressed = cars_tab_btn
                            ElseIf panel_indicator = "CARS" Then
                                ButtonPressed = rest_tab_btn
                            ElseIf panel_indicator = "REST" Then
                                ButtonPressed = unea_tab_btn
                            ElseIf panel_indicator = "UNEA" Then
                                ButtonPressed = jobs_tab_btn
                            ElseIf panel_indicator = "JOBS" Then
                                ButtonPressed = busi_tabs_btn
                            ElseIf panel_indicator = "BUSI" Then
                                ButtonPressed = shel_tab_btn
                            ElseIf panel_indicator = "SHEL" Then
                                ButtonPressed = hst_tab_btn
                            ElseIf panel_indicator = "HEST" Then
                                ButtonPressed = verifs_tab_btn
                            ElseIf panel_indicator = "VERIFS" Then
                                ButtonPressed = notes_tab_btn
                            ElseIf panel_indicator = "NOTES" Then
                                ButtonPressed = finish_btn
                            End If
                        ElseIf panel_array_to_use = UBound(ALL_PANELS_ARRAY, 2) Then
                            ButtonPressed = verifs_tab_btn
                        Else
                            panel_array_to_use = panel_array_to_use + 1
                            Do
								If ALL_PANELS_ARRAY(show_this_panel, panel_array_to_use) = FALSE Then panel_array_to_use = panel_array_to_use + 1
								If panel_array_to_use > UBound(ALL_PANELS_ARRAY, 2) Then
									ButtonPressed = verifs_tab_btn
									panel_array_to_use = UBound(ALL_PANELS_ARRAY, 2)
									Exit Do
								End If

                            Loop until ALL_PANELS_ARRAY(show_this_panel, panel_array_to_use) = TRUE
                            If panel_indicator = ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use) Then find_indicator = FALSE
                            panel_indicator = ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use)
                            second_indicator = ""
                            if panel_indicator = "ACCT" Then second_indicator = "CASH"
                        End If
                    Else
                        panel_indicator = ""
                        second_indicator = ""
                        panel_array_to_use = ""
                    End If

                    If ButtonPressed = address_tab_btn Then
                        panel_indicator = "ADDR"
                        second_indicator = ""
                    ElseIf ButtonPressed = hh_comp_tab_btn Then
                        panel_indicator = "MEMB"
                        second_indicator = ""
                    ElseIf ButtonPressed = wreg_btn Then
                        panel_indicator = "WREG"
                        second_indicator = ""
                    ElseIf ButtonPressed = faci_btn Then
                        panel_indicator = "FACI"
                        second_indicator = ""
                    ElseIf ButtonPressed = revw_btn Then
                        panel_indicator = "REVW"
                        second_indicator = ""
                    ElseIf ButtonPressed = jobs_tab_btn Then
                        panel_indicator = "JOBS"
                        second_indicator = ""
                    ElseIf ButtonPressed = busi_tabs_btn Then
                        panel_indicator = "BUSI"
                        second_indicator = ""
                    ElseIf ButtonPressed = unea_tab_btn Then
                        panel_indicator = "UNEA"
                        second_indicator = ""
                    ElseIf ButtonPressed = child_support_tabs_btn Then
                        panel_indicator = "UNEA"
                        second_indicator = "child support"
                    ElseIf ButtonPressed = acct_tab_btn Then
                        panel_indicator = "ACCT"
                        second_indicator = "CASH"
                    ElseIf ButtonPressed = secu_tab_btn Then
                        panel_indicator = "SECU"
                        second_indicator = ""
                    ElseIf ButtonPressed = cars_tab_btn Then
                        panel_indicator = "CARS"
                        second_indicator = ""
                    ElseIf ButtonPressed = rest_tab_btn Then
                        panel_indicator = "REST"
                        second_indicator = ""
                    ElseIf ButtonPressed = shel_tab_btn Then
                        panel_indicator = "SHEL"
                        second_indicator = ""
                    ElseIf ButtonPressed = hst_tab_btn Then
                        panel_indicator = "HEST"
                        second_indicator = ""
                    ElseIf ButtonPressed = notes_tab_btn Then
                        panel_indicator = "NOTES"
                        second_indicator = ""
                        panel_array_to_use = ""
                    ElseIf ButtonPressed = verifs_tab_btn Then
                        panel_indicator = "VERIFS"
                        second_indicator = ""
                        panel_array_to_use = ""
                    ElseIf ButtonPressed = fill_button Then
                        panel_indicator = "VERIFS"
                        second_indicator = ""
                        panel_array_to_use = ""
                    ElseIf ButtonPressed = finish_btn Then
                        err_msg = ""
                    Else
                    End If

                    ' MsgBox "1st Indicator - " & panel_indicator & vbNewLine & "2nd Indicator - " & second_indicator

                    If ButtonPressed = verifs_tab_btn Then
                        panel_indicator = "VERIFS"
                    ElseIf ButtonPressed = notes_tab_btn Then
                        panel_indicator = "NOTES"
                    ElseIf find_indicator = TRUE AND ButtonPressed <> finish_btn Then
                        For case_panels = 0 to UBound(ALL_PANELS_ARRAY, 2)
                            If ButtonPressed = ALL_PANELS_ARRAY(panel_btn_const, case_panels) Then
                                ' MsgBox "The Button is Pressed"
                                panel_indicator = ALL_PANELS_ARRAY(the_panel_const, case_panels)
                                panel_array_to_use = case_panels
                                ' MsgBox "The panel to use - " & panel_array_to_use & vbNewLine & "The person array indicator - " & ALL_PANELS_ARRAY(array_ref_const, case_panels)
                                Exit For
                            ElseIf ALL_PANELS_ARRAY(the_panel_const, case_panels) = panel_indicator Then
                                panel_array_to_use = case_panels
                                Exit For

                            ElseIf ALL_PANELS_ARRAY(the_panel_const, case_panels) = second_indicator Then
                                panel_array_to_use = case_panels
                                Exit For


                                ' ElseIf second_indicator = "other" Then
                                '     the_array_ref = ALL_PANELS_ARRAY(array_ref_const, case_panels)
                                '     If ALL_INCOME_ARRAY(type_const, the_array_ref) <> "08 - Direct Child Support" AND ALL_INCOME_ARRAY(type_const, the_array_ref) <> "35 - Direct Spousal Support" AND ALL_INCOME_ARRAY(type_const, the_array_ref) <> "36 - Disbursed Child Support" AND ALL_INCOME_ARRAY(type_const, the_array_ref) <> "37 - Disbursed Spousal Support" AND ALL_INCOME_ARRAY(type_const, the_array_ref) <> "39 - Disbursed CS Arrears" AND ALL_INCOME_ARRAY(type_const, the_array_ref) <> "40 - Disbursed Spsl Sup Arrears" Then
                                '         ALL_PANELS_ARRAY(show_this_panel, case_panels) = TRUE
                                '         panel_array_to_use = case_panels
                                '         Exit For
                                '     End If
                                ' ElseIf second_indicator = "child support" Then
                                '     the_array_ref = ALL_PANELS_ARRAY(array_ref_const, case_panels)
                                '     If ALL_INCOME_ARRAY(type_const, the_array_ref) = "08 - Direct Child Support" OR ALL_INCOME_ARRAY(type_const, the_array_ref) = "35 - Direct Spousal Support" OR ALL_INCOME_ARRAY(type_const, the_array_ref) = "36 - Disbursed Child Support" OR ALL_INCOME_ARRAY(type_const, the_array_ref) = "37 - Disbursed Spousal Support" OR ALL_INCOME_ARRAY(type_const, the_array_ref) = "39 - Disbursed CS Arrears" OR ALL_INCOME_ARRAY(type_const, the_array_ref) = "40 - Disbursed Spsl Sup Arrears" Then
                                '         ALL_PANELS_ARRAY(show_this_panel, case_panels) = TRUE
                                '         panel_array_to_use = case_panels
                                '         Exit For
                                '     End If
                                ' End If
                            End If
                        Next
                    End If
                    ' MsgBox "Indicator - " & panel_indicator & vbNewLine & "array place - " & panel_array_to_use & vbNewLine & "Panel in Array - " & ALL_PANELS_ARRAY(the_panel_const, panel_array_to_use)

                    'TODO - look to see about moving REVW status to the end of the script run instead of the beginning to have the script inform report status
                    For faci_panel = 0 to UBound(FACILITIES_ARRAY, 2)
                        If FACILITIES_ARRAY(faci_verif_checkbox, faci_panel) = checked AND FACILITIES_ARRAY(faci_verif_added, faci_panel) <> TRUE Then
                            verifs_needed = verifs_needed & "Verification of Member " & FACILITIES_ARRAY(faci_ref_numb, faci_to_use) & " - " & FACILITIES_ARRAY(faci_member, faci_to_use) & " at facility - " & FACILITIES_ARRAY(faci_name, faci_to_use) & ".; "
                            FACILITIES_ARRAY(faci_verif_added, faci_panel) = TRUE
                        End If
                    Next

                    For asset_panel = 0 to UBound(ALL_ASSETS_ARRAY, 2)
                        If ALL_ASSETS_ARRAY(verif_checkbox_const, asset_panel) = checked AND ALL_ASSETS_ARRAY(verif_added_const, asset_panel) <> TRUE Then
                            If ALL_ASSETS_ARRAY(type_const, asset_panel) = "ACCT" OR ALL_ASSETS_ARRAY(type_const, asset_panel) = "SECU" Then
                                verifs_needed = verifs_needed & "Verif of account at " & ALL_ASSETS_ARRAY(name_const, asset_panel) & " for " & ALL_ASSETS_ARRAY(owner_name, asset_panel) & " during " & ALL_ASSETS_ARRAY(verif_time_const, asset_panel) & ".; "
                            ElseIf ALL_ASSETS_ARRAY(type_const, asset_panel) = "CARS" Then
                                verifs_needed = verifs_needed & "Verif of vehicle " & ALL_ASSETS_ARRAY(year_const, asset_panel) & " " & ALL_ASSETS_ARRAY(make_const, asset_panel) & " " & ALL_ASSETS_ARRAY(model_const, asset_panel) & " belonging to " & ALL_ASSETS_ARRAY(owner_name, asset_panel) & ".; "
                            ElseIf ALL_ASSETS_ARRAY(type_const, asset_panel) = "REST" Then
                                verifs_needed = verifs_needed & "Verif of property (" & right(ALL_ASSETS_ARRAY(type_const, asset_panel), len(ALL_ASSETS_ARRAY(type_const, asset_panel)) - 4) & ") owned by " & ALL_ASSETS_ARRAY(owner_name, asset_panel) & ".; "
                            End If
                            ALL_ASSETS_ARRAY(verif_added_const, asset_panel) = TRUE
                        End If
                    Next

                    For income_panel = 0 to UBound(ALL_INCOME_ARRAY, 2)
                        If ALL_INCOME_ARRAY(verif_checkbox_const, income_panel) = checked AND ALL_INCOME_ARRAY(verif_added_const, income_panel) <> TRUE Then
                            If ALL_INCOME_ARRAY(type_const, income_panel) = "JOBS" Then
                                verifs_needed = verifs_needed & "Verif of job for " & ALL_INCOME_ARRAY(owner_name, income_panel) & " at " & ALL_INCOME_ARRAY(name_const, income_panel) & " (employer) during " & ALL_INCOME_ARRAY(verif_time_const, income_panel) & ".; "
                            ElseIf ALL_INCOME_ARRAY(type_const, income_panel) = "BUSI" Then
                                verifs_needed = verifs_needed & "Verif of self-employment by " & ALL_INCOME_ARRAY(owner_name, income_panel) & " during " & ALL_INCOME_ARRAY(verif_time_const, income_panel) &  ".; "
                            ElseIf ALL_INCOME_ARRAY(type_const, income_panel) = "UNEA" Then
                                verifs_needed = verifs_needed & "Verif of income for " & ALL_INCOME_ARRAY(owner_name, income_panel) & " from " & right(ALL_INCOME_ARRAY(type_const, income_panel), len(ALL_INCOME_ARRAY(type_const, income_panel)) - 5) & " during " & ALL_INCOME_ARRAY(verif_time_const, income_panel) & ".; "
                            End If
                            ALL_INCOME_ARRAY(verif_added_const, asset_panel) = TRUE
                        End If
                    Next

                    For case_pers = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
                        If ALL_CLIENTS_ARRAY(shel_verif_checkbox, case_pers) = checked AND ALL_CLIENTS_ARRAY(shel_verif_added, case_pers) Then
                            verifs_needed = verifs_needed & "Verif of shelter expense for " & ALL_CLIENTS_ARRAY(memb_first_name, case_pers) & " " & ALL_CLIENTS_ARRAY(memb_last_name, case_pers) & " during " & ALL_CLIENTS_ARRAY(shel_verif_time, case_pers) & ".; "

                            ALL_CLIENTS_ARRAY(shel_verif_added, case_pers) = TRUE
                        End If
                    Next

                    If HEST_verif_checkbox = checked and HEST_verif_added <> TRUE Then
                        verifs_needed = verifs_needed & "Detail about utility expenses paid by the household."

                        HEST_verif_added = TRUE
                    End If

                    If ButtonPressed = finish_btn then
                        err_msg = ""

                        If residence_address_match_yn = "No - there is a difference." Then
                            new_resi_one = trim(new_resi_one)
                            new_resi_city = trim(new_resi_city)
                            new_resi_zip = trim(new_resi_zip)

                            If trim(new_addr_effective_date) <> "" AND IsDate(new_addr_effective_date) = FALSE Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, effective date is not a date. Review the effective date."
                            If new_resi_one = "" OR new_resi_city = "" OR new_resi_zip = "" OR new_resi_county = "Select One..." OR new_resi_state = "Select One..." OR new_shel_verif = "Select One..." Then
                                err_msg = err_msg & vbNewLine & "* On ADDR Tab, residence address information is incomplete: "
                                If new_resi_one = "" Then err_msg = err_msg & vbNewLine & "   - Enter house number, street, and apartment number."
                                If new_resi_city = "" Then err_msg = err_msg & vbNewLine & "   - Enter the city."
                                If new_resi_state = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Select the state."
                                If new_resi_zip = "" Then err_msg = err_msg & vbNewLine & "   - Enter the zip code."
                                If new_resi_county = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Enter the county code."
                                If new_shel_verif = "Select One..." Then err_msg = err_msg & vbNewLine &"   - Select the verification information."
                            End If
                        End If
                        If mailing_address_match_yn = "No - there is a difference." Then
                            new_mail_one = trim(new_mail_one)
                            new_mail_city = trim(new_mail_city)
                            new_mail_zip = trim(new_mail_zip)

                            If new_mail_one = "" OR new_mail_city = "" OR new_mail_zip = "" OR new_mail_state = "Select One..." Then
                                err_msg = err_msg & vbNewLine & "* On ADDR Tab, Mailing address information is incomplete: "
                                If new_mail_one = "" Then err_msg = err_msg & vbNewLine & "   - Enter house number, street, and apartment number."
                                If new_mail_city = "" Then err_msg = err_msg & vbNewLine & "   - Enter the city."
                                If new_mail_state = "Select One..." Then err_msg = err_msg & vbNewLine & "   - Select the state."
                                If new_mail_zip = "" Then err_msg = err_msg & vbNewLine & "   - Enter the zip code"
                            End If
                        End If
                        new_phone_one = trim(new_phone_one)
                        new_phone_two = trim(new_phone_two)
                        new_phone_three = trim(new_phone_three)
                        If new_phone_one <> "" Then
                            phone_check = new_phone_one
                            phone_check = replace(phone_check, "(", "")
                            phone_check = replace(phone_check, ")", "")
                            phone_check = replace(phone_check, "-", "")
                            phone_check = replace(phone_check, " ", "")
                            If len(phone_check) <> 10 Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, First listed new phone number does not have 10 numeric digits. Review New Phone Number one."
                            If IsNumeric(phone_check) = FALSE Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, First listed new phone number appears incorrect. Review New Phone Number one."
                        End If
                        If new_phone_two <> "" Then
                            phone_check = new_phone_two
                            phone_check = replace(phone_check, "(", "")
                            phone_check = replace(phone_check, ")", "")
                            phone_check = replace(phone_check, "-", "")
                            phone_check = replace(phone_check, " ", "")
                            If len(phone_check) <> 10 Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, Second listed new phone number does not have 10 numeric digits. Review New Phone Number two."
                            If IsNumeric(phone_check) = FALSE Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, Second listed new phone number appears incorrect. Review New Phone Number two."
                        End If
                        If new_phone_three <> "" Then
                            phone_check = new_phone_three
                            phone_check = replace(phone_check, "(", "")
                            phone_check = replace(phone_check, ")", "")
                            phone_check = replace(phone_check, "-", "")
                            phone_check = replace(phone_check, " ", "")
                            If len(phone_check) <> 10 Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, Third listed new phone number does not have 10 numeric digits. Review New Phone Number three."
                            If IsNumeric(phone_check) = FALSE Then err_msg = err_msg & vbNewLine & "* On ADDR Tab, Third listed new phone number appears incorrect. Review New Phone Number three."
                        End If
                        If verif_reviewed = FALSE Then
                            err_msg = err_msg & vbNewLine & "* To complete the review of panels, the VERIFICATIONS panel should be reviewed."
                            panel_indicator = "VERIFS"
                            second_indicator = ""
                        End If
                        If notes_reviewed = FALSE Then
                            err_msg = err_msg & vbNewLine & "* To complete the review of panels, the NOTES panel should be reviewed."
                            panel_indicator = "NOTES"
                            second_indicator = ""
                        End If
                        If residence_address_match_yn = "No - there is a difference." Then
                            If new_resi_one = "" OR new_resi_city = "" OR new_resi_zip = "" OR new_resi_county = "Select One..." OR new_resi_state = "Select One..." OR new_shel_verif = "Select One..." Then
                                panel_indicator = "ADDR"
                                second_indicator = ""
                            End If
                        End If
                        If mailing_address_match_yn = "No - there is a difference." Then
                            If new_mail_one = "" OR new_mail_city = "" OR new_mail_zip = "" OR new_mail_state = "Select One..." Then
                                panel_indicator = "ADDR"
                                second_indicator = ""
                            End If
                        End If
                        If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
                    End If

                Loop until err_msg = ""
                panel_indicator = "NOTES"
                form_note_reviewed = FALSE
                stat_note_reviewed = FALSE
                verif_note_reviewed = FALSE
                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE
        End If

        full_form_note = ""
        full_form_note_array = ""

        full_form_note = full_form_note & csr_form_date & " CSR Form Information" & "~#~$~"
        full_form_note = full_form_note & "CSR Form Received on: " & csr_form_date
        If client_signed_yn = "Yes" Then full_form_note = full_form_note & " Form signed and "
        If client_signed_yn = "No" Then full_form_note = full_form_note & " Form NOT signed and "
        If client_dated_yn = "Yes" Then full_form_note = full_form_note & "form dated." & "~#~$~"
        If client_dated_yn = "No" Then full_form_note = full_form_note & "form NOT dated." & "~#~$~"
        full_form_note = full_form_note & "--- Answers on form ---" & "~#~$~"
        full_form_note = full_form_note & "Q1. Name and address: " & "~#~$~"
        full_form_note = full_form_note & "    -Client listed Q1 on the form - " & client_on_csr_form & "~#~$~"
        If residence_address_match_yn = "RESI Address not Provided" Then
            full_form_note = full_form_note & "    -Residence address NOT provided on the form." & "~#~$~"
        ElseIf residence_address_match_yn = "No - there is a difference." Then
            full_form_note = full_form_note & "    -Residence address on form differs from the known address." & "~#~$~"
            full_form_note = full_form_note & "     New Residence: " & new_resi_one & "~#~$~"
            full_form_note = full_form_note & "                    " & new_resi_city & ", " & left(new_resi_state, 2) & " " & new_resi_zip & "~#~$~"
            full_form_note = full_form_note & "     New Residence Verif: " & new_shel_verif & "~#~$~"
        ElseIf residence_address_match_yn = "Yes - the addresses are the same." Then
            full_form_note = full_form_note & "    -Residence address on form is the same as the known address." & "~#~$~"
        End If
        If mailing_address_match_yn = "MAIL Address not Provided" Then
            If mail_line_one = "" Then full_form_note = full_form_note & "    -No mailing address known or provided on form." & "~#~$~"
            If mail_line_one <> "" Then full_form_note = full_form_note & "    -Mailing address not provided on form, but known mailing address exists." & "~#~$~"
        ElseIf mailing_address_match_yn = "No - there is a difference." Then
            full_form_note = full_form_note & "    -Mailing address on form differs from the known address." & "~#~$~"
            full_form_note = full_form_note & "     New Mail: " & new_mail_one & "~#~$~"
            full_form_note = full_form_note & "               " & new_mail_city & ", " & left(new_mail_state, 2) & " " & new_mail_zip & "~#~$~"
        ElseIf mailing_address_match_yn = "Yes - the addresses are the same." Then
            full_form_note = full_form_note & "    -Mailing address on form is the same as the known address." & "~#~$~"
        End If
        full_form_note = full_form_note & "    -Homeless: " & homeless_status & "~#~$~"
        If question_one_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_one_notes & "~#~$~"

        full_form_note = full_form_note & "Q2. Has anyone moved in or out of your home? " & quest_two_move_in_out & "~#~$~"
        For known_memb = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
            If ALL_CLIENTS_ARRAY(memb_remo_checkbox, known_memb) = checked Then full_form_note = full_form_note & "    -MOVED OUT - Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_first_name, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb) & "~#~$~"
            If ALL_CLIENTS_ARRAY(memb_new_checkbox, known_memb) = checked Then full_form_note = full_form_note & "    -MOVED IN - Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_first_name, known_memb) & " " & ALL_CLIENTS_ARRAY(memb_last_name, known_memb) & "~#~$~"
        Next
        If new_hh_memb_not_in_mx_yn = "Yes - add another member" Then
            full_form_note = full_form_note & "    -- Household members reported and not listed in STAT --" & "~#~$~"
            For new_hh_memb = 0 to UBound(NEW_MEMBERS_ARRAY, 2)
                If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter) = checked Then
                    full_form_note = full_form_note & "    -MOVED OUT - " & NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb)
                ElseIf NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter) = checked Then
                    full_form_note = full_form_note & "    -MOVED IN - " & NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb)
                Else
                    full_form_note = full_form_note & "    -" & NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb)
                End If
                If trim(NEW_MEMBERS_ARRAY(new_dob, new_hh_memb)) <> "" Then full_form_note = full_form_note & " DOB: " & NEW_MEMBERS_ARRAY(new_dob, new_hh_memb)
                If NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_hh_memb) <> "Select One..." Then full_form_note = full_form_note & ". Rel to applicant: " & NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_hh_memb)
                full_form_note = full_form_note & "~#~$~"
                If NEW_MEMBERS_ARRAY(new_ma_request, new_hh_memb) = checked OR NEW_MEMBERS_ARRAY(new_fs_request, new_hh_memb) = checked OR NEW_MEMBERS_ARRAY(new_grh_request, new_hh_memb) Then
                    full_form_note = full_form_note & "       Member requesting: "
                    If NEW_MEMBERS_ARRAY(new_ma_request, new_hh_memb) = checked Then full_form_note = full_form_note & "HC "
                    If right(full_form_note, 2) <> ": " Then full_form_note = full_form_note & "and "
                    If NEW_MEMBERS_ARRAY(new_fs_request, new_hh_memb) = checked Then full_form_note = full_form_note & "SNAP "
                    If right(full_form_note, 2) <> ": " Then full_form_note = full_form_note & "and "
                    If NEW_MEMBERS_ARRAY(new_grh_request, new_hh_memb) = checked Then full_form_note = full_form_note & "GRH "

                    full_form_note = full_form_note & "~#~$~"
                End If
            Next
        End If
        If question_two_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_two_notes & "~#~$~"

        full_form_note = full_form_note & "-- MA Questions --" & "~#~$~"

        full_form_note = full_form_note & "Q4. Do you want to apply for MA for some who is not getting coverage? " & apply_for_ma & "~#~$~"
        If apply_for_ma = "Yes" Then
            If q_4_details_blank_checkbox = checked Then
                full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
            Else
                For new_ma_client = 0 to UBound(NEW_MA_REQUEST_ARRAY, 2)
                    full_form_note = full_form_note & "    -" & NEW_MA_REQUEST_ARRAY(ma_request_client, new_ma_client) & "~#~$~"
                Next
            End If
        End If
        If question_four_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_four_notes & "~#~$~"

        full_form_note = full_form_note & "Q5. Is anyone self-employed? " & ma_self_employed & "~#~$~"
        If ma_self_employed = "Yes" Then
            If q_5_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_busi = TRUE
        For each_busi = 0 to UBound(NEW_EARNED_ARRAY, 2)
            If NEW_EARNED_ARRAY(earned_type, each_busi) = "BUSI" AND NEW_EARNED_ARRAY(earned_prog_list, each_busi) = "MA" Then
                If first_busi = TRUE then
                    full_form_note = full_form_note & "     --BUSI Information listed on CSR Form--" & "~#~$~"
                    first_busi = FALSE
                End If
                full_form_note = full_form_note & "    -" & NEW_EARNED_ARRAY(earned_client, each_busi) & " Source: " & NEW_EARNED_ARRAY(earned_source, each_busi) & "~#~$~"
                full_form_note = full_form_note & "     Started: " & NEW_EARNED_ARRAY(earned_start_date, each_busi) & ". Yearly Income: $" & NEW_EARNED_ARRAY(earned_amount, each_busi) & "~#~$~"
            End If
        Next
        If question_five_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_five_notes & "~#~$~"

        full_form_note = full_form_note & "Q6. Does anyone work? " & ma_start_working & "~#~$~"
        If ma_start_working = "Yes" Then
            If q_6_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_job = TRUE
        For each_job = 0 to UBound(NEW_EARNED_ARRAY, 2)
            If NEW_EARNED_ARRAY(earned_type, each_job) = "JOBS" AND NEW_EARNED_ARRAY(earned_prog_list, each_job) = "MA" Then
                If first_job = TRUE Then
                    full_form_note = full_form_note & "     --JOBS Information listed on CSR Form--" & "~#~$~"
                    first_job = FALSE
                End If
                If NEW_EARNED_ARRAY(earned_freq, each_job) <> "" Then
                    the_frequency = right(NEW_EARNED_ARRAY(earned_freq, each_job), len(NEW_EARNED_ARRAY(earned_freq, each_job))-4)
                Else
                    the_frequency = "Unknown"
                End If
                full_form_note = full_form_note & "    -" & NEW_EARNED_ARRAY(earned_client, each_job) & " Employer: " & NEW_EARNED_ARRAY(earned_source, each_job) & " Seasonal: " & NEW_EARNED_ARRAY(earned_seasonal, each_job) & "~#~$~"
                full_form_note = full_form_note & "     Started: " & NEW_EARNED_ARRAY(earned_start_date, each_job) & ". Income: $" & NEW_EARNED_ARRAY(earned_amount, each_job) & " " & the_frequency & "~#~$~"
            End If
        Next
        If question_six_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_six_notes & "~#~$~"

        full_form_note = full_form_note & "Q7. Does anyone have unearned income? " & ma_other_income & "~#~$~"
        If ma_other_income = "Yes" Then
            If q_7_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_unea = TRUE
        For each_unea = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
            If NEW_UNEARNED_ARRAY(unearned_type, each_unea) = "UNEA" AND NEW_UNEARNED_ARRAY(unearned_prog_list, each_unea) = "MA" Then
                If first_unea = TRUE Then
                    full_form_note = full_form_note & "     --UNEA Information listed on CSR Form" & "~#~$~"
                    first_unea = FALSE
                End If
                full_form_note = full_form_note & "    -" & NEW_UNEARNED_ARRAY(unearned_client, each_unea) & " Source: " & NEW_UNEARNED_ARRAY(unearned_source, each_unea) & "~#~$~"
                full_form_note = full_form_note & "     Started: " & NEW_UNEARNED_ARRAY(unearned_start_date, each_unea) & " Income: " & NEW_UNEARNED_ARRAY(unearned_amount, each_unea) & " " & NEW_UNEARNED_ARRAY(unearned_freq, each_unea) & "~#~$~"
            End If
        Next
        If question_seven_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_seven_notes & "~#~$~"

        full_form_note = full_form_note & "Q9. Does anyone have liquid assets? " & ma_liquid_assets & "~#~$~"
        If ma_liquid_assets = "Yes" Then
            If q_9_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_account = TRUE
        For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
          If (NEW_ASSET_ARRAY(asset_type, each_asset) = "ACCT" OR NEW_ASSET_ARRAY(asset_type, each_asset) = "CASH") AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
            If first_account = TRUE Then
                full_form_note = full_form_note & "     --Liquid Asset Information listed on CSR Form" & "~#~$~"
                first_account = FALSE
            End If
            full_form_note = full_form_note & "    -" & NEW_ASSET_ARRAY(asset_client, each_asset) & " " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " &  NEW_ASSET_ARRAY(asset_bank_name, each_asset) & "~#~$~"
          End If
        Next
        If question_nine_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_nine_notes & "~#~$~"

        full_form_note = full_form_note & "Q10. Does anyone have any securities? " & ma_security_assets & "~#~$~"
        If ma_security_assets = "Yes" Then
            If q_10_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_secu = TRUE
        For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
          If NEW_ASSET_ARRAY(asset_type, each_asset) = "SECU" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
              If first_secu = TRUE Then
                  full_form_note = full_form_note & "     --Securities Information listed on CSR Form" & "~#~$~"
                  first_secu = FALSE
              End If
              full_form_note = full_form_note & "    -" & NEW_ASSET_ARRAY(asset_client, each_asset) & " " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " &  NEW_ASSET_ARRAY(asset_bank_name, each_asset) & "~#~$~"
          End If
        Next
        If question_ten_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_ten_notes & "~#~$~"

        full_form_note = full_form_note & "Q11. Does anyone own a vehicle? " & ma_vehicle & "~#~$~"
        If ma_vehicle = "Yes" Then
            If q_11_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_car = TRUE
        For each_asset = 0 to UBound(NEW_ASSET_ARRAY, 2)
            If NEW_ASSET_ARRAY(asset_type, each_asset) = "CARS" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
                If first_car = TRUE Then
                    full_form_note = full_form_note & "     --Vehicle Information listed on CSR Form" & "~#~$~"
                    first_car = FALSE
                End If
                full_form_note = full_form_note & "    -" & NEW_ASSET_ARRAY(asset_client, each_asset) & " - " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " - " & NEW_ASSET_ARRAY(asset_year_make_model, each_asset) & "~#~$~"
            End If
        Next
        If question_eleven_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_eleven_notes & "~#~$~"

        full_form_note = full_form_note & "Q12. Does anyone own real estate? " & ma_real_assets & "~#~$~"
        If ma_real_assets = "Yes" Then
            If q_12_details_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        first_home = TRUE
        For each_asset = 0 to Ubound(NEW_ASSET_ARRAY, 2)
            If NEW_ASSET_ARRAY(asset_type, each_asset) = "REST" AND NEW_ASSET_ARRAY(asset_prog_list, each_asset) = "MA" Then
                If first_home = TRUE Then
                    full_form_note = full_form_note & "     --Real Estate Information listed on CSR Form" & "~#~$~"
                    first_home = FALSE
                End If
                full_form_note = full_form_note & "    -" & NEW_ASSET_ARRAY(asset_client, each_asset) & " " & NEW_ASSET_ARRAY(asset_acct_type, each_asset) & " at " & NEW_ASSET_ARRAY(asset_address, each_asset) & "~#~$~"
                y_pos = y_pos + 20
            End If
        Next
        If question_twelve_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_twelve_notes & "~#~$~"

        full_form_note = full_form_note & "Q13. Changes to report? " & ma_other_changes & "~#~$~"
        If ma_other_changes = "Yes" Then
            If changes_reported_blank_checkbox = checked Then full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        End If
        If trim(other_changes_reported) <> "" Then full_form_note = full_form_note & "    -" & other_changes_reported & "~#~$~"

        full_form_note = full_form_note & "-- SNAP Questions --" & "~#~$~"

        full_form_note = full_form_note & "Q15. Has your household moved? " & quest_fifteen_form_answer & "~#~$~"
        If trim(new_rent_or_mortgage_amount) <> "" Then
            full_form_note = full_form_note & "    -New shelter expense amount: $" & new_rent_or_mortgage_amount & "~#~$~"
        Else
            full_form_note = full_form_note & "    -No shelter expense amount listed." & "~#~$~"
        End If
        util_checked = ""
        If heat_ac_checkbox = checked Then util_checked = util_checked & " & Heat/AC"
        If electricity_checkbox = checked Then util_checked = util_checked & " & Electric"
        If telephone_checkbox = checked Then util_checked = util_checked & " & Phone"
        If util_checked = "" Then
            util_checked = "None"
        Else
            util_checked = right(util_checked, len(util_checked) - 3)
        End If
        If question_fifteen_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_fifteen_notes & "~#~$~"

        full_form_note = full_form_note & "Q16. Has there been a change in work income? " & quest_sixteen_form_answer & "~#~$~"
        If q_16_details_blank_checkbox = checked Then
            full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        Else
            first_earned = TRUE
            For the_earned = 0 to UBound(NEW_EARNED_ARRAY, 2)
                If NEW_EARNED_ARRAY(earned_prog_list, the_earned) = "SNAP" Then
                    If first_earned = TRUE Then
                        full_form_note = full_form_note & "     --Earned Income Change Information listed on CSR Form" & "~#~$~"
                        first_earned = FALSE
                    End If
                    full_form_note = full_form_note & "    -" & NEW_EARNED_ARRAY(earned_client, the_earned) & " from " & NEW_EARNED_ARRAY(earned_source, the_earned) & ". Change on: " & NEW_EARNED_ARRAY(earned_change_date, the_earned) & "~#~$~"
                    full_form_note = full_form_note & "     Amount: $" & NEW_EARNED_ARRAY(earned_amount, the_earned) & " " & NEW_EARNED_ARRAY(earned_freq, the_earned) & ". Hours: " & NEW_EARNED_ARRAY(earned_hours, the_earned) & "~#~$~"
                End If
            Next
        End If
        If question_sixteen_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_sixteen_notes & "~#~$~"

        full_form_note = full_form_note & "Q17. Has there been a change of more than $50 in unearned income? " & quest_seventeen_form_answer & "~#~$~"
        If q_17_details_blank_checkbox = checked Then
            full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        Else
            first_unearned = TRUE
            For the_unearned = 0 to UBound(NEW_UNEARNED_ARRAY, 2)
                If NEW_UNEARNED_ARRAY(unearned_prog_list, the_unearned) = "SNAP" Then
                    If first_unearned = TRUE Then
                        full_form_note = full_form_note & "     --UNEA Income Change Information listed on CSR Form" & "~#~$~"
                        first_unearned = FALSE
                    End If
                    full_form_note = full_form_note & "    -Memb " & NEW_UNEARNED_ARRAY(unearned_client, the_unearned) & " from " & NEW_UNEARNED_ARRAY(unearned_source, the_unearned) & ". Change on: " & NEW_UNEARNED_ARRAY(unearned_change_date, the_unearned) & "~#~$~"
                    full_form_note = full_form_note & "     Amount: $" & NEW_UNEARNED_ARRAY(unearned_amount, the_unearned) & " " & NEW_UNEARNED_ARRAY(unearned_freq, the_unearned) & "~#~$~"
                End If
            Next
        End If
        If question_seventeen_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_seventeen_notes & "~#~$~"

        full_form_note = full_form_note & "Q18. Has anyone had a change in shild support? " & quest_eighteen_form_answer & "~#~$~"
        If q_18_details_blank_checkbox = checked Then
            full_form_note = full_form_note & "    -No detail provided" & "~#~$~"
        Else
            first_cs = TRUE
            For the_cs = 0 to UBound(NEW_CHILD_SUPPORT_ARRAY, 2)
                If NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) <> "" OR NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs) <> "" OR NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) <> "Select..." Then
                    If first_cs = TRUE Then
                        full_form_note = full_form_note & "     --Child Support Income Change Information listed on CSR Form" & "~#~$~"
                        first_cs = FALSE
                    End If
                    full_form_note = full_form_note & "    -CS Paid by: " & NEW_CHILD_SUPPORT_ARRAY(cs_payer, the_cs) & ". Amount: $" & NEW_CHILD_SUPPORT_ARRAY(cs_amount, the_cs) & "~#~$~"
                    full_form_note = full_form_note & "     Currently Paying: " & NEW_CHILD_SUPPORT_ARRAY(cs_current, the_cs) & "~#~$~"
                End If
            Next
        End If
        If question_eighteen_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_eighteen_notes & "~#~$~"

        full_form_note = full_form_note & "Q19. Did you work 20 hours per week? " & quest_nineteen_form_answer & "~#~$~"
        If question_nineteen_notes <> "" Then full_form_note = full_form_note & "    -Notes: " & question_nineteen_notes & "~#~$~"

        full_form_note = full_form_note & "---" & "~#~$~"
        full_form_note = full_form_note & worker_signature

        ' https://edocs.dhs.state.mn.us/lfserver/Public/DHS-5576-ENG

        full_stat_note = ""
        If snap_sr_yn = "Yes" Then full_stat_note = full_stat_note & snap_sr_mo & "/" & snap_sr_yr & " SNAP"
        If hc_sr_yn = "Yes" Then
            If full_stat_note = "" Then full_stat_note = full_stat_note & hc_sr_mo & "/" & hc_sr_yr & " HC"
            If full_stat_note <> "" Then full_stat_note = full_stat_note & " & " & hc_sr_mo & "/" & hc_sr_yr & " HC"
        End If
        If grh_sr_yn = "Yes" Then
            If full_stat_note = "" Then full_stat_note = full_stat_note & grh_sr_mo & "/" & grh_sr_yr & " GRH"
            If full_stat_note <> "" Then full_stat_note = full_stat_note & " & " & grh_sr_mo & "/" & grh_sr_yr & " GRH"
        End If

        full_stat_note = full_stat_note & " CSR Info and Processing" & "~#~$~"

        full_stat_note = full_stat_note & "===== HOUSEHOLD COMPOSITION =====" & "~#~$~"
        If residence_address_match_yn = "RESI Address not Provided" Then
            full_stat_note = full_stat_note & "Address was not reported on the CSR form and could not be reviewed." & "~#~$~"
        Else
            full_stat_note = full_stat_note & "Address Reviewed: "
            If residence_address_match_yn = "No - there is a difference." Then full_stat_note = full_stat_note & "A new address has been reported on the CSR." & "~#~$~"
            If residence_address_match_yn = "Yes - the addresses are the same." Then full_stat_note = full_stat_note & "The address reported on the CSR is that same as the known address." & "~#~$~"
        End If
        If mailing_address_match_yn = "MAIL Address not Provided" Then
            If mail_line_one = "" Then
                full_stat_note = full_stat_note & "Case has no reported MAILING address." & "~#~$~"
            Else
                full_stat_note = full_stat_note & "MAILING Address was not reported on the CSR form and could not be reviewed." & "~#~$~"
            End If
        Else
            full_stat_note = full_stat_note & "MAILING Address Reviewed: "
            If mailing_address_match_yn = "No - there is a difference." Then full_stat_note = full_stat_note & "A new address has been reported on the CSR." & "~#~$~"
            If mailing_address_match_yn = "Yes - the addresses are the same." Then full_stat_note = full_stat_note & "The address reported on the CSR is that same as the known address." & "~#~$~"
        End If
        If update_addr_checkbox = checked Then full_stat_note = full_stat_note & " - ADDR updated with new information." & "~#~$~"

        If snap_sr_yn = "Yes" Then
            full_stat_note = full_stat_note & "Members on SNAP: " & snap_active_count & "~#~$~"     'DETAIL TO ADD - SNAP HH Count/info'
            If snap_active_count <> 0 Then
                full_stat_note = full_stat_note & "   - Members listed active: "
                For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
                    If right(full_stat_note, 2) <> ": " Then full_stat_note = full_stat_note & ", "
                    If ALL_CLIENTS_ARRAY(clt_snap_status, known_memb) = "Active" Then full_stat_note = full_stat_note & "Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb)
                Next
                full_stat_note = full_stat_note & "~#~$~"
            End If
        End If
        If hc_sr_yn = "Yes" Then
            full_stat_note = full_stat_note & "Members on HC: " & hc_active_count & "~#~$~"     'DETAIL TO ADD - SNAP HH Count/info'
            If hc_active_count <> 0 Then
                full_stat_note = full_stat_note & "   - Members listed active: "
                For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
                    If right(full_stat_note, 2) <> ": " Then full_stat_note = full_stat_note & ", "
                    If ALL_CLIENTS_ARRAY(clt_hc_status, known_memb) = "Active" Then full_stat_note = full_stat_note & "Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb)
                Next
                full_stat_note = full_stat_note & "~#~$~"
            End If
        End If
        If grh_sr_yn = "Yes" Then
            full_stat_note = full_stat_note & "Members on GRH: " & grh_active_count & "~#~$~"     'DETAIL TO ADD - SNAP HH Count/info'
            If grh_active_count <> 0 Then
                full_stat_note = full_stat_note & "   - Members listed active: "
                For known_memb = 0 to UBOUND(ALL_CLIENTS_ARRAY, 2)
                    If right(full_stat_note, 2) <> ": " Then full_stat_note = full_stat_note & ", "
                    If ALL_CLIENTS_ARRAY(clt_grh_status, known_memb) = "Active" Then full_stat_note = full_stat_note & "Memb " & ALL_CLIENTS_ARRAY(memb_ref_numb, known_memb)
                Next
                full_stat_note = full_stat_note & "~#~$~"
            End If
        End If

        If new_hh_memb_not_in_mx_yn = "Yes - add another member" Then
            full_stat_note = full_stat_note & "-- Members reported and not yet added to MAXIS --" & "~#~$~"

        End If

        If new_hh_memb_not_in_mx_yn = "Yes - add another member" Then
            full_stat_note = full_stat_note & "-- Household members reported and not listed in STAT --" & "~#~$~"
            For new_hh_memb = 0 to UBound(NEW_MEMBERS_ARRAY, 2)
                If NEW_MEMBERS_ARRAY(new_memb_moved_out, new_memb_counter) = checked Then
                    full_stat_note = full_stat_note & "     - MOVED OUT - " & NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb)
                ElseIf NEW_MEMBERS_ARRAY(new_memb_moved_in, new_memb_counter) = checked Then
                    full_stat_note = full_stat_note & "     - MOVED IN - " & NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb)
                Else
                    full_stat_note = full_stat_note & "     - " & NEW_MEMBERS_ARRAY(new_full_name, new_hh_memb)
                End If
                If trim(NEW_MEMBERS_ARRAY(new_dob, new_hh_memb)) <> "" Then full_stat_note = full_stat_note & " DOB: " & NEW_MEMBERS_ARRAY(new_dob, new_hh_memb)
                If NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_hh_memb) <> "Select One..." Then full_stat_note = full_stat_note & ". Rel to applicant: " & NEW_MEMBERS_ARRAY(new_rel_to_applicant, new_hh_memb)
                full_stat_note = full_stat_note & "~#~$~"
                If NEW_MEMBERS_ARRAY(new_ma_request, new_hh_memb) = checked OR NEW_MEMBERS_ARRAY(new_fs_request, new_hh_memb) = checked OR NEW_MEMBERS_ARRAY(new_grh_request, new_hh_memb) Then
                    full_stat_note = full_stat_note & "       Member requesting: "
                    If NEW_MEMBERS_ARRAY(new_ma_request, new_hh_memb) = checked Then full_stat_note = full_stat_note & "HC "
                    If right(full_stat_note, 2) <> ": " Then full_stat_note = full_stat_note & "and "
                    If NEW_MEMBERS_ARRAY(new_fs_request, new_hh_memb) = checked Then full_stat_note = full_stat_note & "SNAP "
                    If right(full_stat_note, 2) <> ": " Then full_stat_note = full_stat_note & "and "
                    If NEW_MEMBERS_ARRAY(new_grh_request, new_hh_memb) = checked Then full_stat_note = full_stat_note & "GRH "
                    full_stat_note = full_stat_note & "~#~$~"
                End If
            Next
        End If

        If faci_panel_exists = TRUE Then
            full_stat_note = full_stat_note & "===== FACILITIES =====" & "~#~$~"
            For faci_listed = 0 to UBound(FACILITIES_ARRAY, 2)
                full_stat_note = full_stat_note & "* Facility: " & FACILITIES_ARRAY(faci_name, faci_listed) & ", Vendor #: " & FACILITIES_ARRAY(faci_vendor_number, faci_listed) & ", Type: " &  FACILITIES_ARRAY(faci_type, faci_listed) & "~#~$~"
                full_stat_note = full_stat_note & "   - Member in FACI: Member " & FACILITIES_ARRAY(faci_ref_numb, faci_listed) & "~#~$~"
            Next
        End If
        full_stat_note = full_stat_note & "===== ASSETS =====" & "~#~$~"
        asset_found = FALSE
        For the_asset = 0 to UBound(ALL_ASSETS_ARRAY, 2)
            If ALL_ASSETS_ARRAY(category_const, the_asset) = "CASH" Then full_stat_note = full_stat_note & "* M" & ALL_ASSETS_ARRAY(owner_name, the_asset) & " - CASH Amount: $" & ALL_ASSETS_ARRAY(amount_const, the_asset) & "~#~$~"
            If ALL_ASSETS_ARRAY(category_const, the_asset) = "ACCT" OR ALL_ASSETS_ARRAY(category_const, the_asset) = "SECU" Then full_stat_note = full_stat_note & "* M" & ALL_ASSETS_ARRAY(owner_name, the_asset) & " " & right(ALL_ASSETS_ARRAY(type_const, the_asset), len(ALL_ASSETS_ARRAY(type_const, the_asset)) - 3) & " at " & ALL_ASSETS_ARRAY(name_const, the_asset) & ". Balance: $" & ALL_ASSETS_ARRAY(amount_const, the_asset) & "~#~$~"
            If ALL_ASSETS_ARRAY(category_const, the_asset) = "ACCT" OR ALL_ASSETS_ARRAY(category_const, the_asset) = "SECU" OR ALL_ASSETS_ARRAY(category_const, the_asset) = "CARS" Then
                asset_found = TRUE
                If ALL_ASSETS_ARRAY(item_notes_const, the_asset) <> "" Then full_stat_note = full_stat_note & "   - Notes: " & ALL_ASSETS_ARRAY(item_notes_const, the_asset) & "~#~$~"
            End If
        Next
        For the_asset = 0 to UBound(ALL_ASSETS_ARRAY, 2)
            If ALL_ASSETS_ARRAY(category_const, the_asset) = "CARS" Then
                asset_found = TRUE
                full_stat_note = full_stat_note & "* M" & ALL_ASSETS_ARRAY(owner_name, the_asset) & " " & right(ALL_ASSETS_ARRAY(type_const, the_asset), len(ALL_ASSETS_ARRAY(type_const, the_asset)) - 2) & ". " & ALL_ASSETS_ARRAY(make_model_yr, the_asset) & ". " & "~#~$~"
                If ALL_ASSETS_ARRAY(item_notes_const, the_asset) <> "" Then full_stat_note = full_stat_note & "   - Notes: " & ALL_ASSETS_ARRAY(item_notes_const, the_asset) & "~#~$~"
            End If
        Next
        For the_asset = 0 to UBound(ALL_ASSETS_ARRAY, 2)
            If ALL_ASSETS_ARRAY(category_const, the_asset) = "REST" Then
                asset_found = TRUE
                full_stat_note = full_stat_note & "* M" & ALL_ASSETS_ARRAY(owner_name, the_asset) & " " & right(ALL_ASSETS_ARRAY(type_const, the_asset), len(ALL_ASSETS_ARRAY(type_const, the_asset)) - 2) & ". Verif: " & right(ALL_ASSETS_ARRAY(verif_const, the_asset), len(ALL_ASSETS_ARRAY(verif_const, the_asset)) - 5) & ". " & "~#~$~"
                If ALL_ASSETS_ARRAY(item_notes_const, the_asset) <> "" Then full_stat_note = full_stat_note & "   - Notes: " & ALL_ASSETS_ARRAY(item_notes_const, the_asset) & "~#~$~"
            End If
        Next
        If asset_found = FALSE Then full_stat_note = full_stat_note & "  -- No Assets listed in STAT --"

        full_stat_note = full_stat_note & "===== INCOME =====" & "~#~$~"
        income_found = FALSE
        For the_income = 0 to UBound(ALL_INCOME_ARRAY, 2)
            If ALL_INCOME_ARRAY(category_const, the_income) = "BUSI" Then
                income_found = TRUE
                full_stat_note = full_stat_note & "* M" & ALL_INCOME_ARRAY(owner_name, the_income) & " - SELF - " & right(ALL_INCOME_ARRAY(type_const, the_income), len(ALL_INCOME_ARRAY(type_const, the_income)) - 5) & "~#~$~"
                If snap_sr_yn = "Yes" Then
                    full_stat_note = full_stat_note & "   - Gross Income: $" & ALL_INCOME_ARRAY(busi_snap_gross_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_snap_gross_prosp, the_income) & "(prosp). Verif: " & ALL_INCOME_ARRAY(busi_snap_income_verif, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Gross Expense: $" & ALL_INCOME_ARRAY(busi_snap_expense_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_snap_expense_prosp, the_income) & "(prosp). Verif: " & ALL_INCOME_ARRAY(busi_snap_expense_verif, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Net Income: $" & ALL_INCOME_ARRAY(busi_snap_net_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_snap_net_prosp, the_income) & "(prosp)." & "~#~$~"
                End If
                If hc_sr_yr = "Yes" Then
                    full_stat_note = full_stat_note & "   - Gross Income: $" & ALL_INCOME_ARRAY(busi_hc_b_gross_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_hc_b_gross_prosp, the_income) & "(prosp). Verif: " & ALL_INCOME_ARRAY(busi_hc_b_income_verif, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Gross Expense: $" & ALL_INCOME_ARRAY(busi_hc_b_expense_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_hc_b_expense_prosp, the_income) & "(prosp). Verif: " & ALL_INCOME_ARRAY(busi_hc_b_expense_verif, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Net Income: $" & ALL_INCOME_ARRAY(busi_hc_b_net_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_hc_b_net_prosp, the_income) & "(prosp)." & "~#~$~"
                End If
                If grh_sr_yn = "Yes" Then
                    full_stat_note = full_stat_note & "   - Gross Income: $" & ALL_INCOME_ARRAY(busi_cash_gross_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_cash_gross_prosp, the_income) & "(prosp). Verif: " & ALL_INCOME_ARRAY(busi_cash_income_verif, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Gross Expense: $" & ALL_INCOME_ARRAY(busi_cash_expense_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_cash_expense_prosp, the_income) & "(prosp). Verif: " & ALL_INCOME_ARRAY(busi_cash_expense_verif, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Net Income: $" & ALL_INCOME_ARRAY(busi_cash_net_retro, the_income) & "(retro) - $" & ALL_INCOME_ARRAY(busi_cash_net_prosp, the_income) & "(prosp)." & "~#~$~"
                End If
            End If
            If ALL_INCOME_ARRAY(category_const, the_income) = "JOBS" Then
                income_found = TRUE
                full_stat_note = full_stat_note & "* M" & ALL_INCOME_ARRAY(owner_name, the_income) & " at " & ALL_INCOME_ARRAY(name_const, the_income) & ", verification: " & right(ALL_INCOME_ARRAY(verif_const, the_income), len(ALL_INCOME_ARRAY(verif_const, the_income)) - 4) & "~#~$~"
                If hc_sr_yr = "Yes" Then
                    full_stat_note = full_stat_note & "   - Retro Income: $" & ALL_INCOME_ARRAY(retro_income_amount, the_income) & ". Hours: " & ALL_INCOME_ARRAY(retro_income_hours, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Prospective Income: $" & ALL_INCOME_ARRAY(pay_amt_const, the_income) & ". Hours: " & ALL_INCOME_ARRAY(pay_amt_const, the_income) & "~#~$~"
                End If
                If snap_sr_yn = "Yes" Then
                    If ALL_INCOME_ARRAY(snap_pic_monthly_income, the_income) = "" Then
                        full_stat_note = full_stat_note & "   - SNAP PIC has not been updated." & "~#~$~"
                    Else
                        If ALL_INCOME_ARRAY(snap_pic_frequency, the_income) <> "" Then full_stat_note = full_stat_note & "   - SNAP Prospective Monthly Income: $" & ALL_INCOME_ARRAY(snap_pic_monthly_income, the_income) & " Paid: " & right(ALL_INCOME_ARRAY(snap_pic_frequency, the_income), len(ALL_INCOME_ARRAY(snap_pic_frequency, the_income)) - 4) & "~#~$~"
                        If ALL_INCOME_ARRAY(snap_pic_frequency, the_income) = "" Then full_stat_note = full_stat_note & "   - SNAP Prospective Monthly Income: $" & ALL_INCOME_ARRAY(snap_pic_monthly_income, the_income) & " Pay Frequency not known." & "~#~$~"
                        full_stat_note = full_stat_note & "   - Income per pay date: " & ALL_INCOME_ARRAY(snap_pic_income_per_pay, the_income) & ", hours per pay date: " & ALL_INCOME_ARRAY(snap_pic_hours_per_pay, the_income) & "~#~$~"
                    End If
                End If
                If grh_sr_yn = "Yes" Then
                    full_stat_note = full_stat_note & "   - GRH Prospective Monthly Income: $" & ALL_INCOME_ARRAY(grh_pic_monthly_income, the_income) & " Paid: " & right(ALL_INCOME_ARRAY(grh_pic_frequency, the_income), len(ALL_INCOME_ARRAY(grh_pic_frequency, the_income)) - 4) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Income per pay date: " & ALL_INCOME_ARRAY(grh_pic_income_per_pay, the_income) & "~#~$~"
                End If
                If ALL_INCOME_ARRAY(verif_checkbox_const, ei_to_use) = checked Then full_stat_note = full_stat_note & "   - This job needs to be verified." & "~#~$~"
            End If
            If ALL_INCOME_ARRAY(category_const, the_income) = "UNEA" Then
                income_found = TRUE
                full_stat_note = full_stat_note & "* M" & ALL_INCOME_ARRAY(owner_name, the_income) & " income from " & right(ALL_INCOME_ARRAY(type_const, the_income), len(ALL_INCOME_ARRAY(type_const, the_income)) - 5) & ", verification: " & right(ALL_INCOME_ARRAY(verif_const, the_income), len(ALL_INCOME_ARRAY(verif_const, the_income)) - 4) & "~#~$~"
                If hc_sr_yn = "Yes" OR grh_sr_yn = "Yes" Then
                    full_stat_note = full_stat_note & "   - Retro Income: $" & ALL_INCOME_ARRAY(retro_income_amount, the_income) & "~#~$~"
                    full_stat_note = full_stat_note & "   - Prospective Income: $" & ALL_INCOME_ARRAY(amount_const, the_income) & "~#~$~"
                End If
                If snap_sr_yr = "Yes" Then
                    If ALL_INCOME_ARRAY(snap_pic_monthly_income, the_income) = "" Then
                        full_stat_note = full_stat_note & "   - SNAP PIC has not been updated." & "~#~$~"
                    Else
                        full_stat_note = full_stat_note & "   - SNAP Prospective Monthly Income: $" & ALL_INCOME_ARRAY(snap_pic_monthly_income, the_income) & " Paid: " & right(ALL_INCOME_ARRAY(snap_pic_frequency, the_income), len(ALL_INCOME_ARRAY(snap_pic_frequency, the_income)) - 4) & "~#~$~"
                        full_stat_note = full_stat_note & "   - Income per pay date: " & ALL_INCOME_ARRAY(snap_pic_income_per_pay, the_income) & "~#~$~"
                    End If
                End If
                If ALL_INCOME_ARRAY(verif_checkbox_const, ei_to_use) = checked Then full_stat_note = full_stat_note & "   - This income source needs to be verified." & "~#~$~"
            End If
        Next

        If income_found = FALSE Then full_stat_note = full_stat_note & "  -- No Income listed in STAT --" & "~#~$~"

        full_stat_note = full_stat_note & "===== EXPENSES =====" & "~#~$~"
        For the_client = 0 to UBound(ALL_CLIENTS_ARRAY, 2)
            If ALL_CLIENTS_ARRAY(shel_paid_to, the_client) <> "" Then
                full_stat_note = full_stat_note & "* M" & ALL_CLIENTS_ARRAY(memb_ref_nbr, the_client) & " " & ALL_CLIENTS_ARRAY(memb_first_name, the_client) & " " & ALL_CLIENTS_ARRAY(memb_last_name, the_client) & " pays sheler expense to " & ALL_CLIENTS_ARRAY(shel_paid_to, the_client) & ". " & "~#~$~"

                If ALL_CLIENTS_ARRAY(shel_rent_prosp_amt, the_client) <> "" Then        full_stat_note = full_stat_note & "   - Rent Amount: $" & ALL_CLIENTS_ARRAY(shel_rent_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_rent_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_rent_retro_verif, the_client)) - 5) & "~#~$~"
                If ALL_CLIENTS_ARRAY(shel_lot_rent_prosp_amt, the_client) <> "" Then    full_stat_note = full_stat_note & "   - Lot Rent Amount: $" & ALL_CLIENTS_ARRAY(shel_lot_rent_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_lot_rent_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_lot_rent_retro_verif, the_client)) - 5) & "~#~$~"
                If ALL_CLIENTS_ARRAY(shel_mortgage_prosp_amt, the_client) <> "" Then    full_stat_note = full_stat_note & "   - Mortgage Amount: $" & ALL_CLIENTS_ARRAY(shel_mortgage_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_mortgage_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_mortgage_retro_verif, the_client)) - 5) & "~#~$~"
                If ALL_CLIENTS_ARRAY(shel_insurance_prosp_amt, the_client) <> "" Then   full_stat_note = full_stat_note & "   - Insurance Amount: $" & ALL_CLIENTS_ARRAY(shel_insurance_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_insurance_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_insurance_retro_verif, the_client)) - 5) & "~#~$~"
                If ALL_CLIENTS_ARRAY(shel_tax_prosp_amt, the_client) <> "" Then         full_stat_note = full_stat_note & "   - Taxes Amount: $" & ALL_CLIENTS_ARRAY(shel_tax_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_tax_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_tax_retro_verif, the_client)) - 5) & "~#~$~"
                If ALL_CLIENTS_ARRAY(shel_room_prosp_amt, the_client) <> "" Then        full_stat_note = full_stat_note & "   - Room Amount: $" & ALL_CLIENTS_ARRAY(shel_room_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_room_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_room_retro_verif, the_client)) - 5) & "~#~$~"
                If ALL_CLIENTS_ARRAY(shel_garage_prosp_amt, the_client) <> "" Then      full_stat_note = full_stat_note & "   - Garage Amount: $" & ALL_CLIENTS_ARRAY(shel_garage_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_garage_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_garage_retro_verif, the_client)) - 5) & "~#~$~"

                ' full_stat_note = full_stat_note & "      "
                If ALL_CLIENTS_ARRAY(shel_hud_sub_yn, the_client) = "Y" Then
                    full_stat_note = full_stat_note & "   - Shelter expense is subsidized."
                    If ALL_CLIENTS_ARRAY(shel_subsidy_prosp_amt, the_client) <> "" Then
                        full_stat_note = full_stat_note & " Subsidy amount: $" & ALL_CLIENTS_ARRAY(shel_subsidy_prosp_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_subsidy_prosp_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_subsidy_prosp_verif, the_client)) - 5) & "~#~$~"
                    ElseIf ALL_CLIENTS_ARRAY(shel_subsidy_retro_amt, the_client) <> "" Then
                        full_stat_note = full_stat_note & " Subsidy amount: $" & ALL_CLIENTS_ARRAY(shel_subsidy_retro_amt, the_client) & ", verif: " & right(ALL_CLIENTS_ARRAY(shel_subsidy_retro_verif, the_client), len(ALL_CLIENTS_ARRAY(shel_subsidy_prosp_verif, the_client)) - 5) & "~#~$~"
                    End If
                End If
            End If
        Next
        If HEST_persons_paying <> "" Then
            full_stat_note = full_stat_note & "* Utilities Expense: Paid by: " & HEST_persons_paying & ". Total Expense: $" & HEST_total_expense & "~#~$~"
            If HEST_retro_heat_air <> "" Then full_stat_note = full_stat_note & "   - Retro Heat/AC: " & HEST_retro_heat_air & " - $ " & HEST_retro_heat_air_amount & "~#~$~"
            If HEST_prosp_heat_air <> "" Then full_stat_note = full_stat_note & "   - Prosp Heat/AC: " & HEST_prosp_heat_air & " - $ " & HEST_prosp_heat_air_amount & "~#~$~"
            If HEST_retro_electric <> "" Then full_stat_note = full_stat_note & "   - Retro Electric: " & HEST_retro_electric & " - $ " & HEST_retro_electric_amount & "~#~$~"
            If HEST_prosp_electric <> "" Then full_stat_note = full_stat_note & "   - Prosp Electric: " & HEST_prosp_electric & " - $ " & HEST_prosp_electric_amount & "~#~$~"
            If HEST_retro_phone <> "" Then full_stat_note = full_stat_note & "   - Retro Phone: " & HEST_retro_phone & " - $ " & HEST_retro_phone_amount & "~#~$~"
            If HEST_prosp_phone <> "" Then full_stat_note = full_stat_note & "   - Prosp Phone: " & HEST_prosp_phone & " - $ " & HEST_prosp_phone_amount & "~#~$~"
        End If
        full_stat_note = full_stat_note & "===== SR PROCESS =====" & "~#~$~"
        If snap_sr_yn = "Yes" Then
            full_stat_note = full_stat_note & "* SNAP SR for " & snap_sr_mo & "/" & snap_sr_yr & ". Status: " & snap_sr_status & "~#~$~"
            If snap_questions_complete = TRUE Then full_stat_note = full_stat_note & "   - SR Form submitted on " & csr_form_date & ", and appears complete by the client for SNAP." & "~#~$~"
            If snap_questions_complete = FALSE Then full_stat_note = full_stat_note & "   - SR Form submitted on " & csr_form_date & ", and appears INCOMPLETE by the client for SNAP." & "~#~$~"
            full_stat_note = full_stat_note & "   - Verifications for information provided is " & SNAP_verifs_needed & "~#~$~"
            If snap_sr_notes <> "" Then full_stat_note = full_stat_note & "   - NOTES: " & snap_sr_notes & "~#~$~"
        End If
        If hc_sr_yn = "Yes" Then
            full_stat_note = full_stat_note & "* HC SR for " & hc_sr_mo & "/" & hc_sr_yr & ". Status: " & hc_sr_status & "~#~$~"
            If ma_questions_complete = TRUE Then full_stat_note = full_stat_note & "   - SR Form submitted on " & csr_form_date & ", and appears complete by the client for HC." & "~#~$~"
            If ma_questions_complete = FALSE Then full_stat_note = full_stat_note & "   - SR Form submitted on " & csr_form_date & ", and appears INCOMPLETE by the client for HC." & "~#~$~"
            full_stat_note = full_stat_note & "   - Verifications for information provided is " & HC_verifs_needed & "~#~$~"
            If hc_sr_notes <> "" Then full_stat_note = full_stat_note & "   - NOTES: " & hc_sr_notes & "~#~$~"
        End If
        If grh_sr_yn = "Yes" Then
            full_stat_note = full_stat_note & "* HC SR for " & grh_sr_mo & "/" & grh_sr_yr & ". Status: " & grh_sr_status & "~#~$~"
            If grh_questions_complete = TRUE Then full_stat_note = full_stat_note & "   - SR Form submitted on " & csr_form_date & ", and appears complete by the client for GRH." & "~#~$~"
            If grh_questions_complete = FALSE Then full_stat_note = full_stat_note & "   - SR Form submitted on " & csr_form_date & ", and appears INCOMPLETE by the client for GRH." & "~#~$~"
            full_stat_note = full_stat_note & "   - Verifications for information provided is " & GRH_verifs_needed & "~#~$~"
            If grh_sr_notes <> "" Then full_stat_note = full_stat_note & "   - NOTES: " & grh_sr_notes & "~#~$~"
        End If
        If HRF_checkbox = checked then full_stat_note = full_stat_note & "* CSR and cash supplement used as HRF." & "~#~$~"
        If eDRS_sent_checkbox = checked then full_stat_note = full_stat_note & "* eDRS sent." & "~#~$~"
        IF Sent_arep_checkbox = checked THEN full_stat_note = full_stat_note & "* Sent form(s) to AREP." & "~#~$~"
        If MADE_checkbox = checked then full_stat_note = full_stat_note & "* Emailed MADE through DHS-SIR." & "~#~$~"
        full_stat_note = full_stat_note & "* ACTIONS TAKEN: " &  actions_taken & "~#~$~"
        full_stat_note = full_stat_note & "* NOTES: " &  other_notes & "~#~$~"
        full_stat_note = full_stat_note & "---" & "~#~$~"
        full_stat_note = full_stat_note & worker_signature


        full_verif_note = ""

        If verifs_needed <> "" Then
            verif_counter = 1
            verifs_array = ""
            verifs_needed = trim(verifs_needed)
            If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
            If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
            If InStr(verifs_needed, ";") <> 0 Then
                verifs_array = split(verifs_needed, ";")
            Else
                verifs_array = array(verifs_needed)
            End If

            full_verif_note = full_verif_note & "VERIFICATIONS REQUESTED" & vbCr
            If verif_req_form_sent_date <> "" Then
                full_verif_note = full_verif_note & "* Verif request form sent on: " & verif_req_form_sent_date & vbCr
                full_verif_note = full_verif_note & "---" & vbCr
            End If
            full_verif_note = full_verif_note & "List of all verifications requested:" & vbCr
            For each verif_item in verifs_array
                verif_item = trim(verif_item)
                If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
                verif_counter = verif_counter + 1
                full_verif_note = full_verif_note & verif_item & vbCr
            Next
            full_verif_note = full_verif_note & "---" & vbCr
            full_verif_note = full_verif_note & worker_signature & vbCr
            ' full_verif_note = full_verif_note & "" & vbCr
            ' full_verif_note = full_verif_note & "" & vbCr

        End If

        full_form_note_array = split(full_form_note, "~#~$~")
        ' form_note_block_one = ""
        ' form_note_block_two = ""
        ' form_note_block_three = ""
        ' line_count = 0
        ' box_count = 1
        ' For each note_line in full_form_note_array
        '   If box_count = 1 Then form_note_block_one = form_note_block_one & note_line & vbCr
        '   If box_count = 2 Then form_note_block_two = form_note_block_two & note_line & vbCr
        '   If box_count = 3 Then form_note_block_three = form_note_block_three & note_line & vbCr
        '   line_count = line_count + 1
        '   If line_count = 35 Then
        '       line_count = 1
        '       box_count = box_count + 1
        '   End If
        ' Next

        full_stat_note_array = split(full_stat_note, "~#~$~")
        ' stat_note_block_one = ""
        ' stat_note_block_two = ""
        ' stat_note_block_three = ""
        ' line_count = 0
        ' box_count = 1
        ' For each note_line in full_stat_note_array
        '     If box_count = 1 Then stat_note_block_one = stat_note_block_one & note_line & vbCr
        '     If box_count = 2 Then stat_note_block_two = stat_note_block_two & note_line & vbCr
        '     If box_count = 3 Then stat_note_block_three = stat_note_block_three & note_line & vbCr
        '     line_count = line_count + 1
        '     If line_count = 35 then
        '         line_count = 1
        '         box_count = box_count + 1
        '     End If
        ' Next

        ' MsgBox full_form_note
        ' MsgBox full_stat_note
        ' MsgBox "form box 1" & vbNewLine & form_note_block_one
        ' MsgBox "form box 2" & vbNewLine & form_note_block_two
        ' MsgBox "form box 3" & vbNewLine & form_note_block_three
        '
        ' MsgBox "stat box 1" & vbNewLine & stat_note_block_one
        ' MsgBox "stat box 2" & vbNewLine & stat_note_block_two
        ' MsgBox "stat box 3" & vbNewLine & stat_note_block_three
        '
        ' form_note_block_one = "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        ' form_note_block_two = "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        ' form_note_block_three = "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        ' stat_note_block_one = "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        ' stat_note_block_two = "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        ' stat_note_block_three = "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '
        ' For x = 0 to 40
        '     form_note_block_one = form_note_block_one & x & "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '     form_note_block_two = form_note_block_two & x & "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '     form_note_block_three = form_note_block_three & x & "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '     stat_note_block_one = stat_note_block_one & x & "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '     stat_note_block_two = stat_note_block_two & x & "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '     stat_note_block_three = stat_note_block_three & x & "XXXXXXXXXXXXXXXXXXXXXXXX" & vbCr
        '     ' MsgBox x
        ' Next
        ' MsgBox "form box 1" & vbNewLine & form_note_block_one

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 830, 380, "Confirm CASE:NOTE Information"
          GroupBox 590, 305, 230, 73, "Confirm CASE/NOTE Veribiage"
          Text 600, 315, 190, 30, "This is the wording that will be entered into CASE/NOTE. Review thoroughly here.                                                         Are these notes accurate and complete?"
          ' Text 600, 330, 230, 10, "Are these notes accurate and complete?"
          If full_verif_note = "" Then Text 440, 345, 95, 15, "  Verification CASE/NOTE              -No verifs indicated-"

          If note_to_show = form_note Then
              form_note_reviewed = TRUE
              Text 375, 10, 300, 10, "CASE NOTE about CSR Form"

              ' Text 5, 20, 260, 280, form_note_block_one
              ' Text 280, 30, 260, 270, form_note_block_two
              ' Text 555, 30, 260, 260,form_note_block_three

              x_pos = 5
              y_pos = 20
              For each note_line in full_form_note_array
                If len(note_line) > 90 Then
                    Text x_pos, y_pos, 260, 20, note_line
                    y_pos = y_pos + 10
                Else
                    Text x_pos, y_pos, 260, 10, note_line
                End If
                y_pos = y_pos + 10
                If y_pos = 340 Then
                    y_pos = 30
                    x_pos = x_pos + 275
                End If
              Next

              Text 305, 348, 50, 10, "Form Detail"
              ButtonGroup ButtonPressed
                ' PushButton 10, 325, 70, 15, "Form Detail", form_note_btn
                PushButton 360, 345, 70, 15, "STAT Info", stat_note_btn
                If full_verif_note <> "" Then PushButton 430, 345, 95, 15, "Verification CASE/NOTE", verif_note_btn
          End If
          If note_to_show = stat_note Then
              stat_note_reviewed = TRUE
              Text 375, 10, 300, 10, "CASE NOTE about STAT/CASE Information"
              '
              ' Text 5, 20, 270, 295, stat_note_block_one
              ' Text 280, 30, 270, 285, stat_note_block_two
              ' Text 555, 30, 270, 285, stat_note_block_three

              x_pos = 5
              y_pos = 20
              For each note_line in full_stat_note_array
                If len(note_line) > 90 Then
                    Text x_pos, y_pos, 260, 20, note_line
                    y_pos = y_pos + 10
                Else
                    Text x_pos, y_pos, 260, 10, note_line
                End If
                y_pos = y_pos + 10
                If y_pos = 340 Then
                    y_pos = 30
                    x_pos = x_pos + 275
                End If
              Next

              ' Text 10, 20, 555, 275, full_stat_note
              Text 375, 348, 55, 10, "STAT Info"
              ButtonGroup ButtonPressed
                PushButton 290, 345, 70, 15, "Form Detail", form_note_btn
                ' PushButton 80, 325, 70, 15, "STAT Info", stat_note_btn
                If full_verif_note <> "" Then PushButton 430, 345, 95, 15, "Verification CASE/NOTE", verif_note_btn
          End If
          If note_to_show = verif_note Then
              verif_note_reviewed = TRUE
              Text 375, 10, 300, 10, "Verification CASE NOTE"
              Text 10, 20, 555, 275, full_verif_note
              Text 435, 348, 97, 10, "Verification CASE/NOTE"
              ButtonGroup ButtonPressed
                PushButton 290, 345, 70, 15, "Form Detail", form_note_btn
                PushButton 360, 345, 70, 15, "STAT Info", stat_note_btn
                ' If full_verif_note <> "" Then PushButton 150, 325, 95, 15, "Verification CASE/NOTE", verif_note_btn
          End If
          GroupBox 5, 342, 550, 2, ""
          ButtonGroup ButtonPressed
            PushButton 600, 345, 200, 15, "No - this information is not accurate/complete.", case_note_NOT_correct_btn
            PushButton 600, 360, 200, 15, "Yes - thiese CASE/NOTE(s) look accurate and complete", case_not_is_all_good_btn

          Text 15, 345, 270, 50, "Since this script operates in a new manner and has altered functionality, this is the final step in the entering of information. Reviewing this information is paramount in case action accuracy and effective communication. "
        EndDialog

        err_msg = ""

        dialog Dialog1
        cancel_confirmation

        If ButtonPressed = case_note_NOT_correct_btn Then
            show_all_the_panels = TRUE
            panel_indicator = "NOTES"
        End If

        If ButtonPressed = case_not_is_all_good_btn Then
            If full_verif_note <> "" AND verif_note_reviewed = FALSE Then
                err_msg = err_msg & vbNewLine & "* Review the CASE/NOTE for the Verification Information. All of the notes must be reviewed before indicating the notes are accurate and complete."
                ButtonPressed = verif_note_btn
            End If
            If stat_note_reviewed = FALSE THen
                err_msg = err_msg & vbNewLine & "* Review the CASE/NOTE for the STAT Information. All of the notes must be reviewed before indicating the notes are accurate and complete."
                ButtonPressed = stat_note_btn
            End If
            If form_note_reviewed = FALSE Then
                err_msg = err_msg & vbNewLine & "* Review the CASE/NOTE for the Form Information. All of the notes must be reviewed before indicating the notes are accurate and complete."
                ButtonPressed = form_note_btn
            End If
        End If
        If ButtonPressed = form_note_btn OR ButtonPressed = verif_note_btn OR ButtonPressed = stat_note_btn Then show_all_the_panels = FALSE
        If ButtonPressed = form_note_btn Then note_to_show = form_note
        If ButtonPressed = stat_note_btn Then note_to_show = stat_note
        If ButtonPressed = verif_note_btn Then note_to_show = verif_note

        If err_msg <> "" Then MsgBox "Please review all notes before continuing:" & vbNewLine & err_msg
    Loop until ButtonPressed = case_not_is_all_good_btn
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE
'CASE NOTES

If update_addr_checkbox = checked Then
    If IsDate(new_addr_effective_date) <> TRUE Then new_addr_effective_date = date
    original_footer_month = MAXIS_footer_month
    original_footer_year = MAXIS_footer_year

    the_month = DatePart("m", new_addr_effective_date)
    MAXIS_footer_month = right("00" & the_month, 2)
    the_year = DatePart("yyyy", new_addr_effective_date)
    MAXIS_footer_year = right(the_year, 2)

    Call back_to_SELF
	Call access_ADDR_panel("WRITE", notes_on_address, new_resi_one, "", resi_street_full, new_resi_city, new_resi_state, new_resi_zip, "", "", "", "", living_situation_status, reservation_name, new_mail_one, "", mail_street_full, new_mail_city, new_mail_state, new_mail_zip, new_addr_effective_date, "", new_phone_one, new_phone_two, new_phone_three, phone_one_type, phone_two_type, phone_three_type, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

    If IsDate(new_addr_effective_date) = TRUE Then
        MAXIS_footer_month = original_footer_month
        MAXIS_footer_year = original_footer_year
    End If

End If

create_word_document = FALSE

If export_form_info_to_word_checkbox = checked Then create_word_document = TRUE
If export_verifs_info_to_work_checkbox = checked Then create_word_document = TRUE

If create_word_document = TRUE Then
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    Set objDoc = objWord.Documents.Add()
    objWord.Caption = "Information for Verification Request"
    Set objSelection = objWord.Selection
    objSelection.PageSetup.LeftMargin = 50
    objSelection.PageSetup.RightMargin = 50
    objSelection.PageSetup.TopMargin = 30
    objSelection.PageSetup.BottomMargin = 25
    objSelection.ParagraphFormat.SpaceAfter = 0
    objSelection.Font.Name = "Ariel"

    If export_form_info_to_word_checkbox = checked Then
        objSelection.Font.Size = "14"
        objSelection.Font.Bold = TRUE
        objSelection.TypeText "Form Information" & vbCr & vbCr

        objSelection.Font.Size = "11"
        objSelection.Font.Italic = TRUE
        If form_questions_complete = FALSE Then objSelection.TypeText "CSR Form - all items were NOT answered." & vbCr
        If form_questions_complete = TRUE Then objSelection.TypeText "CSR Form - all items were answered." & vbCr

        objSelection.Font.Bold = FALSE
        objSelection.Font.Italic = FALSE

        If q_one_complete = FALSE Then
            If client_on_csr_form = "Person Information Missing" Then
                objSelection.TypeText "Q1 - name info." & vbCr
            End If
            If residence_address_match_yn = "RESI Address not Provided" Then
                objSelection.TypeText "Q1 - address info." & vbCr
            End If
        End If
        If q_two_complete = FALSE Then
            If quest_two_move_in_out = "Did not answer" Then
                objSelection.TypeText "Q2 - answer yes or no." & vbCr
            ElseIf quest_two_move_in_out = "Yes" AND hh_memb_change = FALSE THen
                objSelection.TypeText "Q2 - answer was Yes but no information about a member moving in or out blank." & vbCr
            End If
        End If


        If HC_active = TRUE Then
            If q_five_complete = FALSE Then
                If ma_self_employed = "Did not answer" Then
                    objSelection.TypeText "Q5 - answer yes or no." & vbCr
                ElseIf ma_self_employed = "Yes" AND q_5_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q5 - answer was Yes but Self Employment information blank." & vbCr
                End If
            End If
            If q_six_complete = FALSE Then
                If ma_start_working = "Did not answer" Then
                    objSelection.TypeText "Q6 - answer yes or no." & vbCr
                ElseIf ma_start_working = "Yes" AND q_6_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q6 - answer was Yes but work information blank." & vbCr
                End If
            End If
            If q_seven_complete = FALSE Then
                If ma_other_income = "Did not_answer" Then
                    objSelection.TypeText "Q7 - answer yes or no." & vbCr
                ElseIf ma_other_income = "Yes" AND q_7_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q7 - answer was Yes but  other income information blank." & vbCr
                End If
            End If
            If q_nine_complete = FALSE Then
                If ma_liquid_assets = "Did not answer" Then
                    objSelection.TypeText "Q9 - answer yes or no." & vbCr
                ElseIf ma_liquid_assets = "Yes" AND q_9_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q9 - answer was Yes but account information blank." & vbCr
                End If
            End If
            If q_ten_complete = FALSE Then
                If ma_security_assets = "Did not answer" Then
                    objSelection.TypeText "Q10 - answer yes or no." & vbCr
                ElseIf ma_security_assets = "Yes" and q_10_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q10 - answer was Yes but other asset information blank." & vbCr
                End If
            End If
            If q_eleven_complete = FALSE Then
                If ma_vehicle = "Did not answer" Then
                    objSelection.TypeText "Q11 - answer yes or no." & vbCr
                ElseIf ma_vehicle = "Yes" AND q_11_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q11 - answer was Yes but vehicle detail blank." & vbCr
                End If
            End If
            If q_twelve_complete = FALSE Then
                If ma_real_assets = "Did not answer" Then
                    objSelection.TypeText "Q12 - answer yes or no." & vbCr
                ElseIf ma_real_assets = "Yes" and q_12_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q12 - answer was Yes but real estate information blank." & vbCr
                End If
            End If
            If q_thirteen_complete = FALSE Then
                If ma_other_changes = "Did not answer" Then
                    objSelection.TypeText "Q13 - answer yes or no." & vbCr
                ElseIf ma_other_changes = "Yes" and changes_reported_blank_checkbox = checked Then
                    objSelection.TypeText "Q13 - Answer was Yes but no information listed." & vbCr
                End If
            End If
        End If

        If SNAP_active = TRUE Then
            If q_fifteen_complete = FALSE Then
                objSelection.TypeText "Q15 - answer yes or no." & vbCr
            End If
            If q_sixteen_complete = FALSE Then
                If quest_sixteen_form_answer = "Did not answer" Then
                    objSelection.TypeText "Q16 - Answer was Yes but no change of income information blank." & vbCr
                ElseIf quest_sixteen_form_answer = "Yes" AND q_16_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q16 - answer yes or no." & vbCr
                End If
            End If
            If q_seventeen_complete = FALSE Then
                If quest_seventeen_form_answer = "Did not answer" Then
                    objSelection.TypeText "Q17 - answer was Yes other income information blank." & vbCr
                ElseIf quest_seventeen_form_answer = "Yes" AND q_17_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q17 - answer yes or no." & vbCr
                End If
            End If
            If q_eightneen_complete = FALSE Then
                If quest_eighteen_form_answer = "Did not answer" Then
                    objSelection.TypeText "Q18 - answer yes or no." & vbCr
                ElseIf quest_eighteen_form_answer = "Yes" and q_18_details_blank_checkbox = checked Then
                    objSelection.TypeText "Q18 - answer was Yes but child support information blank." & vbCr
                End If
            End If
            If q_nineteen_complete = FALSE Then
                objSelection.TypeText "Q19 - answer yes or no." & vbCr
            End If
        End If
    End If

    If export_verifs_info_to_work_checkbox = checked Then
        objSelection.Font.Size = "14"
        objSelection.Font.Bold = TRUE
        objSelection.TypeText "Verifications Needed" & vbCr & vbCr

        objSelection.Font.Size = "11"
        objSelection.Font.Italic = TRUE
        If verifs_needed = "" Then objSelection.TypeText "No Verifications Listed" & vbCr
        If verifs_needed <> "" Then objSelection.TypeText "Verifications Listed:" & vbCr
        If verif_req_form_sent_date <> "" Then objSelection.TypeText "Verif request form sent on: " & verif_req_form_sent_date & vbCr

        objSelection.Font.Bold = FALSE
        objSelection.Font.Italic = FALSE

        verif_counter = 1
        For each verif_item in verifs_array
            verif_item = trim(verif_item)
            If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
            verif_counter = verif_counter + 1
            objSelection.TypeText verif_item
        Next
    End If
End If

If full_verif_note <> "" Then
    verif_counter = 1
    call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("VERIFICATIONS REQUESTED")
    If verif_req_form_sent_date <> "" Then
        Call write_variable_in_CASE_NOTE("* Verif request form sent on: " & verif_req_form_sent_date)
        Call write_variable_in_CASE_NOTE("---")
    End If
    Call write_variable_in_CASE_NOTE("List of all verifications requested:")
    For each verif_item in verifs_array
        verif_item = trim(verif_item)
        If number_verifs_checkbox = checked Then
            verif_item = verif_counter & ". " & verif_item
        Else
            verif_item = "- " & verif_item
        End If
        verif_counter = verif_counter + 1
        Call write_variable_in_CASE_NOTE("  " & verif_item)
    Next
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)
End If

call start_a_blank_CASE_NOTE

For each note_line in full_form_note_array
    call write_variable_in_CASE_NOTE(note_line)
Next
PF3

call start_a_blank_CASE_NOTE

For each note_line in full_stat_note_array
    call write_variable_in_CASE_NOTE(note_line)
Next


'ASK ALL THE SNAP QUESIONS

'HAVE THE SCRIPT ASSESS THE INFORMATION PROVIDED ON THE FORM TO DETERMINE IF COMPLETE OR INCOMPLETE AND RELAY THAT TO THE WORKER.

'SERIES OF DIALOGS FROM THE QUESTIONS TO CLARIFY AND DETAIL

'SCRIPT WILL GUIDE THROUGH THE VERIFS NEEDED AND THE ACTIONS REQUIRED

'Case note and any TIKLS etc'
script_end_procedure_with_error_report("Success! CSR Information was reviewed and noted. Please make sure to accept the Work items in ECF associated with this CSR. Thank you!")
