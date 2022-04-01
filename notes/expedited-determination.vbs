'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EXPEDITED DETERMINATION.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("09/01/2021", "Expedited Determination Functionality has been completely enhanced.##~####~##The functionality to guide through the assesment of a case meeting expedited criteria has been updated. This new functionality adds a series of 3 new dialogs to support this process.##~####~##This new functionality matches the scripts NOTES - Expedited Determination and the new script NOTES - Interview.##~##", "Casey Love, Hennepin County")
call changelog_update("03/05/2020", "Added enhanced handling for the month the script will use to look at information. The best informaiton is provided in the month of application.", "Casey Love, Hennepin County")
call changelog_update("05/28/2019", "Updates to read the Expedited Screening case note.", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS------------------------------------------------------------------------------------------------------------------
Function format_explanation_text(text_variable)
	text_variable = trim(text_variable)
	Do while Instr(text_variable, "; ;") <> 0
		text_variable = replace(text_variable, "; ;", "; ")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, ";;") <> 0
		text_variable = replace(text_variable, ";;", "; ")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, "  ") <> 0
		text_variable = replace(text_variable, "  ", " ")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, "  ") <> 0
		text_variable = replace(text_variable, ".; .", "")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, "  ") <> 0
		text_variable = replace(text_variable, "; .;", "")
		text_variable = trim(text_variable)
	Loop
	Do while left(text_variable, 1) = "."
		text_variable = right(text_variable, len(text_variable) - 1)
		text_variable = trim(text_variable)
		Do while left(text_variable, 1) = ";"
			text_variable = right(text_variable, len(text_variable) - 1)
			text_variable = trim(text_variable)
		Loop
	Loop
	Do while left(text_variable, 1) = ";"
		text_variable = right(text_variable, len(text_variable) - 1)
		text_variable = trim(text_variable)
	Loop
	Do while right(text_variable, 1) = ";"
		text_variable = left(text_variable, len(text_variable) - 1)
		text_variable = trim(text_variable)
	Loop
	text_variable = trim(text_variable)
End Function

function app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, JOBS_ARRAY, BUSI_ARRAY, UNEA_ARRAY)
	return_btn = 5001
	enter_btn = 5002
	add_another_jobs_btn = 5005
	remove_one_jobs_btn = 5006
	add_another_busi_btn = 5007
	remove_one_busi_btn = 5008
	add_another_unea_btn = 5009
	remove_one_unea_btn = 2010
	income_review_completed = True

	original_income = determined_income
	determined_income = 0
	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 296, 160, "Determination of Assets in Month of Application"
		  DropListBox 210, 40, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", jobs_income_yn
		  DropListBox 210, 60, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", busi_income_yn
		  DropListBox 235, 110, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", unea_income_yn
		  ButtonGroup ButtonPressed
		    PushButton 240, 140, 50, 15, "Enter", enter_btn
		  Text 10, 10, 205, 10, "Does this household have any income?"
		  GroupBox 10, 25, 255, 65, "Earned Income "
		  Text 65, 45, 140, 10, "Is anyone in the household working a job?"
		  Text 25, 65, 180, 10, "Does anyone in the household have self employment?"
		  GroupBox 10, 95, 280, 40, "Unearned Income"
		  Text 20, 115, 215, 10, "Does anyone in the household receive any other kind of income?"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then
			income_review_completed = False
			Exit Do
		End If

		If jobs_income_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has Income from a Job."
		If busi_income_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has Income from Self Employment."
		If unea_income_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has Income from Another Source."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If income_review_completed = True Then
		Do
			prvt_err_msg = ""

			If jobs_income_yn = "No" Then jobs_grp_len = 30
			If jobs_income_yn = "Yes" Then jobs_grp_len = 55 + (UBound(JOBS_ARRAY, 2) + 1) * 20
			If busi_income_yn = "No" Then busi_grp_len = 30
			If busi_income_yn = "Yes" Then busi_grp_len = 55 + (UBound(BUSI_ARRAY, 2) + 1) * 20
			If unea_income_yn = "No" Then unea_grp_len = 30
			If unea_income_yn = "Yes" Then unea_grp_len = 55 + (UBound(UNEA_ARRAY, 2) + 1) * 20

			dlg_len = 45 + jobs_grp_len + busi_grp_len + unea_grp_len

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 400, dlg_len, "Determination of Income in Month of Application"
			  ' Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
			  ButtonGroup ButtonPressed
				  y_pos = 10
				  GroupBox 10, y_pos, 380, jobs_grp_len, "JOBS"
				  y_pos = y_pos + 15
				  If jobs_income_yn = "Yes" Then
					  Text 20, y_pos, 190, 10, "JOBS Income on this case"
					  y_pos = y_pos + 15
					  Text 20, y_pos, 50, 10, "Employee"
					  Text 90, y_pos, 70, 10, "Employer/Job"
					  Text 185, y_pos, 50, 10, "Hourly Wage"
					  Text 245, y_pos, 50, 10, "Weekly Hours"
					  Text 305, y_pos, 50, 10, "Pay Frequency"
					  y_pos = y_pos + 10

					  For the_job = 0 to UBound(JOBS_ARRAY, 2)
					  	  JOBS_ARRAY(jobs_wage_const, the_job) = JOBS_ARRAY(jobs_wage_const, the_job) & ""
						  JOBS_ARRAY(jobs_hours_const, the_job) = JOBS_ARRAY(jobs_hours_const, the_job) & ""
						  EditBox 20, y_pos, 60, 15, JOBS_ARRAY(jobs_employee_const, the_job)
						  EditBox 90, y_pos, 85, 15, JOBS_ARRAY(jobs_employer_const, the_job)
						  EditBox 185, y_pos, 50, 15, JOBS_ARRAY(jobs_wage_const, the_job)
						  EditBox 245, y_pos, 50, 15, JOBS_ARRAY(jobs_hours_const, the_job)
						  DropListBox 305, y_pos, 75, 15, "Select One..."+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", JOBS_ARRAY(jobs_frequency_const, the_job)
						  y_pos = y_pos + 20
					  Next
					  PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_jobs_btn
					  PushButton 320, y_pos, 60, 10, "REMOVE ONE", remove_one_jobs_btn
					  y_pos = y_pos + 20
				  Else
					  Text 20, y_pos, 355, 10, "This household does NOT have JOBS."
					  y_pos = y_pos + 20
				  End If

				  GroupBox 10, y_pos, 380, busi_grp_len, "Self Employment"
				  y_pos = y_pos + 15
				  If busi_income_yn = "Yes" Then
					  Text 20, y_pos, 190, 10, "BUSI Income on this case"
					  y_pos = y_pos + 15
					  Text 20, y_pos, 65, 10, "Business Owner"
					  Text 125, y_pos, 70, 10, "Business"
					  Text 230, y_pos, 65, 10, "Monthly Earnings"
					  Text 290, y_pos, 65, 10, "Annual Earnings"
					  y_pos = y_pos + 10
					  ' Text 305, y_pos, 50, 10, "Pay Frequency"
					  For the_busi = 0 to UBound(BUSI_ARRAY, 2)
					  	  BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = BUSI_ARRAY(busi_monthly_earnings_const, the_busi) & ""
						  BUSI_ARRAY(busi_annual_earnings_const, the_busi) = BUSI_ARRAY(busi_annual_earnings_const, the_busi) & ""

						  EditBox 20, y_pos, 95, 15, BUSI_ARRAY(busi_owner_const, the_busi)
						  EditBox 125, y_pos, 95, 15, BUSI_ARRAY(busi_info_const, the_busi)
						  EditBox 230, y_pos, 50, 15, BUSI_ARRAY(busi_monthly_earnings_const, the_busi)
						  EditBox 290, y_pos, 50, 15, BUSI_ARRAY(busi_annual_earnings_const, the_busi)
						  y_pos = y_pos + 20
					  Next
					  PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_busi_btn
					  PushButton 320, y_pos, 60, 10, "REMOVE ONE", remove_one_busi_btn
					  y_pos = y_pos + 20
				  Else
					  Text 20, y_pos, 355, 10, "This household does NOT have BUSI."
					  y_pos = y_pos + 20
				  End If

				  GroupBox 10, y_pos, 380, unea_grp_len, "Unearned"
				  y_pos = y_pos + 15
				  If unea_income_yn = "Yes" Then
					  Text 20, y_pos, 190, 10, "UNEA Income on this case"
					  y_pos = y_pos + 15
					  Text 20, y_pos, 65, 10, "Member Receiving"
					  Text 125, y_pos, 70, 10, "Income Type"
					  Text 230, y_pos, 65, 10, "Monthly Amount"
					  Text 290, y_pos, 65, 10, "Weekly Amount"
					  y_pos = y_pos + 10
					  ' Text 305, y_pos, 50, 10, "Pay Frequency"
					  For the_unea = 0 to UBound(UNEA_ARRAY, 2)
						  UNEA_ARRAY(unea_monthly_earnings_const, the_unea) = UNEA_ARRAY(unea_monthly_earnings_const, the_unea) & ""
						  UNEA_ARRAY(unea_weekly_earnings_const, the_unea) = UNEA_ARRAY(unea_weekly_earnings_const, the_unea) & ""
						  EditBox 20, y_pos, 95, 15, UNEA_ARRAY(unea_owner_const, the_unea)
						  EditBox 125, y_pos, 95, 15, UNEA_ARRAY(unea_info_const, the_unea)
						  EditBox 230, y_pos, 50, 15, UNEA_ARRAY(unea_monthly_earnings_const, the_unea)
						  EditBox 290, y_pos, 50, 15, UNEA_ARRAY(unea_weekly_earnings_const, the_unea)
						  y_pos = y_pos + 20
					  Next
					  PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_unea_btn
					  PushButton 320, y_pos, 60, 10, "REMOVE ONE", remove_one_unea_btn
					  y_pos = y_pos + 20
				  Else
					  Text 20, y_pos, 355, 10, "This household does NOT have UNEA."
					  y_pos = y_pos + 20
				  End If

				  PushButton 345, dlg_len - 20, 50, 15, "Return", return_btn
			EndDialog

			dialog Dialog1

			If ButtonPressed = 0 Then
				income_review_completed = False
				Exit Do
			End If
			If ButtonPressed = -1 Then ButtonPressed = return_btn

			last_jobs_item = UBound(JOBS_ARRAY, 2)
			If ButtonPressed = add_another_jobs_btn Then
				last_jobs_item = last_jobs_item + 1
				ReDim Preserve JOBS_ARRAY(jobs_notes_const, last_jobs_item)
			End If
			If ButtonPressed = remove_one_jobs_btn Then
				last_jobs_item = last_jobs_item - 1
				ReDim Preserve JOBS_ARRAY(jobs_notes_const, last_jobs_item)
			End If

			last_busi_item = UBound(BUSI_ARRAY, 2)
			If ButtonPressed = add_another_busi_btn Then
				last_busi_item = last_busi_item + 1
				ReDim Preserve BUSI_ARRAY(busi_notes_const, last_busi_item)
			End If
			If ButtonPressed = remove_one_unea_btn Then
				last_busi_item = last_busi_item - 1
				ReDim Preserve BUSI_ARRAY(busi_notes_const, last_busi_item)
			End If

			last_unea_item = UBound(UNEA_ARRAY, 2)
			If ButtonPressed = add_another_unea_btn Then
				last_unea_item = last_unea_item + 1
				ReDim Preserve UNEA_ARRAY(unea_notes_const, last_unea_item)
			End If
			If ButtonPressed = remove_one_busi_btn Then
				last_unea_item = last_unea_item - 1
				ReDim Preserve UNEA_ARRAY(unea_notes_const, last_unea_item)
			End If

			For the_job = 0 to UBound(JOBS_ARRAY, 2)
				JOBS_ARRAY(jobs_employee_const, the_job) = trim(JOBS_ARRAY(jobs_employee_const, the_job))
				JOBS_ARRAY(jobs_employer_const, the_job) = trim(JOBS_ARRAY(jobs_employer_const, the_job))
				JOBS_ARRAY(jobs_wage_const, the_job) = trim(JOBS_ARRAY(jobs_wage_const, the_job))
				JOBS_ARRAY(jobs_hours_const, the_job) = trim(JOBS_ARRAY(jobs_hours_const, the_job))
				JOBS_ARRAY(jobs_frequency_const, the_job) = trim(JOBS_ARRAY(jobs_frequency_const, the_job))

				If JOBS_ARRAY(jobs_employee_const, the_job) <> "" OR JOBS_ARRAY(jobs_employer_const, the_job) <> "" OR JOBS_ARRAY(jobs_wage_const, the_job) <> "" OR JOBS_ARRAY(jobs_hours_const, the_job) <> "" Then
					jobs_err_msg = ""
					If JOBS_ARRAY(jobs_employee_const, the_job) = "" Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the name of the employer for this JOB."
					If JOBS_ARRAY(jobs_employer_const, the_job) = "" Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the employer for This JOB."
					If IsNumeric(JOBS_ARRAY(jobs_wage_const, the_job)) = False Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the amount that " & JOBS_ARRAY(jobs_employee_const, the_job) & " is paid per hour from " & JOBS_ARRAY(jobs_employer_const, the_job) & " as a number."
					If IsNumeric(JOBS_ARRAY(jobs_hours_const, the_job)) = False Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the number of hours " & JOBS_ARRAY(jobs_employee_const, the_job) & " works per week in the application month for " & JOBS_ARRAY(jobs_employer_const, the_job) & " as a number."
					If JOBS_ARRAY(jobs_frequency_const, the_job) = "Select One..." Then jobs_err_msg = jobs_err_msg & vbCr & "* Select the pay frequency that " & JOBS_ARRAY(jobs_employee_const, the_job) & " receives their checks in from " & JOBS_ARRAY(jobs_employer_const, the_job) & "."
					If jobs_err_msg <> "" Then prvt_err_msg = prvt_err_msg & vbCr & "For the JOB that is Number " & the_job + 1 & " on the list." & vbCr & jobs_err_msg & vbCr
				End If
			Next

			For the_busi = 0 to UBound(BUSI_ARRAY, 2)
				BUSI_ARRAY(busi_owner_const, the_busi) = trim(BUSI_ARRAY(busi_owner_const, the_busi))
				BUSI_ARRAY(busi_info_const, the_busi) = trim(BUSI_ARRAY(busi_info_const, the_busi))
				BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = trim(BUSI_ARRAY(busi_monthly_earnings_const, the_busi))
				BUSI_ARRAY(busi_annual_earnings_const, the_busi) = trim(BUSI_ARRAY(busi_annual_earnings_const, the_busi))

				If BUSI_ARRAY(busi_owner_const, the_busi) <> "" OR BUSI_ARRAY(busi_info_const, the_busi) <> "" OR BUSI_ARRAY(busi_monthly_earnings_const, the_busi) <> "" OR BUSI_ARRAY(busi_annual_earnings_const, the_busi) <> "" Then
					busi_err_msg = ""
					If BUSI_ARRAY(busi_owner_const, the_busi) = "" Then busi_err_msg = busi_err_msg & vbCr & "* Enter the name of the employer for this Self Employment."
					If BUSI_ARRAY(busi_info_const, the_busi) = "" Then busi_err_msg = busi_err_msg & vbCr & "* Enter the business information for this Self Employment."
					If BUSI_ARRAY(busi_monthly_earnings_const, the_busi) <> "" AND IsNumeric(BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = False Then busi_err_msg = busi_err_msg & vbCr & "* Enter the amount that " & BUSI_ARRAY(busi_owner_const, the_busi) & " earns monthly from " & BUSI_ARRAY(busi_info_const, the_busi) & "."
					If BUSI_ARRAY(busi_annual_earnings_const, the_busi) <> "" AND IsNumeric(BUSI_ARRAY(busi_annual_earnings_const, the_busi)) = False Then busi_err_msg = busi_err_msg & vbCr & "* Enter the number of hours " & BUSI_ARRAY(busi_owner_const, the_busi) & " earns yearly from " & BUSI_ARRAY(busi_info_const, the_busi) & "."
					If IsNumeric(BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = True AND IsNumeric(BUSI_ARRAY(busi_annual_earnings_const, the_busi)) = True Then
						BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = FormatNumber(BUSI_ARRAY(busi_monthly_earnings_const, the_busi), 2, -1, 0, -1)
						BUSI_ARRAY(busi_annual_earnings_const, the_busi) = FormatNumber(BUSI_ARRAY(busi_annual_earnings_const, the_busi), 2, -1, 0, -1)
						annual_from_monthly = 0
						annual_from_monthly =  BUSI_ARRAY(busi_monthly_earnings_const, the_busi) * 12
						annual_from_monthly = FormatNumber(annual_from_monthly, 2, -1, 0, -1)
						If annual_from_monthly <> BUSI_ARRAY(busi_annual_earnings_const, the_busi) Then busi_err_msg = busi_err_msg & vbCr & "* The annual amount does not match up with the monthly amount entered. The Annual earnings should be 12 times the Monthly earnings. You only need to enter one of these amounts."
					ElseIf IsNumeric(BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = True AND BUSI_ARRAY(busi_annual_earnings_const, the_busi) = "" Then
						BUSI_ARRAY(busi_annual_earnings_const, the_busi) = FormatNumber(BUSI_ARRAY(busi_monthly_earnings_const, the_busi)*12, 2, -1, 0, -1)
					ElseIf IsNumeric(BUSI_ARRAY(busi_annual_earnings_const, the_busi)) = True AND BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = "" Then
						BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = FormatNumber(BUSI_ARRAY(busi_annual_earnings_const, the_busi)/12, 2, -1, 0, -1)
					End If
					If busi_err_msg <> "" Then prvt_err_msg = prvt_err_msg & vbCr & "For the BUSI that is Number " & the_busi + 1 & " on the list." & vbCr & busi_err_msg & vbCr
				End If
			Next

			For the_unea = 0 to UBound(UNEA_ARRAY, 2)
				unea_err_msg = ""
				UNEA_ARRAY(unea_owner_const, the_unea) = trim(UNEA_ARRAY(unea_owner_const, the_unea))
				UNEA_ARRAY(unea_info_const, the_unea) = trim(UNEA_ARRAY(unea_info_const, the_unea))
				UNEA_ARRAY(unea_monthly_earnings_const, the_unea) = trim(UNEA_ARRAY(unea_monthly_earnings_const, the_unea))
				UNEA_ARRAY(unea_weekly_earnings_const, the_unea) = trim(UNEA_ARRAY(unea_weekly_earnings_const, the_unea))
				If UNEA_ARRAY(unea_owner_const, the_unea) <> "" OR UNEA_ARRAY(unea_info_const, the_unea) <> "" OR UNEA_ARRAY(unea_monthly_earnings_const, the_unea) <> "" OR UNEA_ARRAY(unea_weekly_earnings_const, the_unea) <> "" Then
					If UNEA_ARRAY(unea_owner_const, the_unea) = "" Then unea_err_msg = unea_err_msg & vbCr & "* Enter the name of the the person who received this Unearned Income."
					If UNEA_ARRAY(unea_info_const, the_unea) = "" Then unea_err_msg = unea_err_msg & vbCr & "* Enter the information of what type of Unearned Income this is listed."
					If IsNumeric(UNEA_ARRAY(unea_monthly_earnings_const, the_unea)) = True AND IsNumeric(UNEA_ARRAY(unea_weekly_earnings_const, the_unea)) = True Then
						If FormatNumber(UNEA_ARRAY(unea_monthly_earnings_const, the_unea), 0) <> FormatNumber(UNEA_ARRAY(unea_weekly_earnings_const, the_unea) * 4.3, 0) Then unea_err_msg = unea_err_msg & vbCr & "* Enter Only one of the following: Weekly Amount or Monthly Amount"
					ElseIf IsNumeric(UNEA_ARRAY(unea_monthly_earnings_const, the_unea)) = False AND UNEA_ARRAY(unea_weekly_earnings_const, the_unea) = "" Then
						unea_err_msg = unea_err_msg & vbCr & "* Enter the amount that " & UNEA_ARRAY(unea_owner_const, the_unea) & " receives per month from " & UNEA_ARRAY(unea_info_const, the_unea) & " as a number."
					ElseIf IsNumeric(UNEA_ARRAY(unea_weekly_earnings_const, the_unea)) = False AND UNEA_ARRAY(unea_monthly_earnings_const, the_unea) = "" Then
						unea_err_msg = unea_err_msg & vbCr & "* Enter the number of hours " & UNEA_ARRAY(unea_owner_const, the_unea) & " receives per week from " & UNEA_ARRAY(unea_info_const, the_unea) & " as a number."
					End IF
					If unea_err_msg <> "" Then prvt_err_msg = prvt_err_msg & vbCr & "For the UNEA that is Number " & the_unea + 1 & " on the list." & vbCr & unea_err_msg & vbCr
				End If
			Next

			If prvt_err_msg <> "" AND ButtonPressed = return_btn Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
		Loop Until ButtonPressed = return_btn AND prvt_err_msg = ""
	End If

	For the_job = 0 to UBound(JOBS_ARRAY, 2)
		If IsNumeric(JOBS_ARRAY(jobs_wage_const, the_job)) = True AND IsNumeric(JOBS_ARRAY(jobs_hours_const, the_job)) = True Then
			weekly_pay = JOBS_ARRAY(jobs_wage_const, the_job) * JOBS_ARRAY(jobs_hours_const, the_job)
			JOBS_ARRAY(jobs_monthly_pay_const, the_job) = weekly_pay * 4.3
			determined_income = determined_income + JOBS_ARRAY(jobs_monthly_pay_const, the_job)
		End If
	Next

	For the_busi = 0 to UBound(BUSI_ARRAY, 2)
		If IsNumeric(BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = True Then determined_income = determined_income + BUSI_ARRAY(busi_monthly_earnings_const, the_busi)
	Next
	For the_unea = 0 to UBound(UNEA_ARRAY, 2)
		If IsNumeric(UNEA_ARRAY(unea_monthly_earnings_const, the_unea)) = True Then
			determined_income = determined_income + UNEA_ARRAY(unea_monthly_earnings_const, the_unea)
		ElseIf IsNumeric(UNEA_ARRAY(unea_weekly_earnings_const, the_unea)) = True Then
			monthly_pay = UNEA_ARRAY(unea_weekly_earnings_const, the_unea) * 4.3
			determined_income = determined_income + monthly_pay
		End If
	Next
	determined_income = FormatNumber(determined_income, 2, -1, 0, -1)

	If income_review_completed = False Then determined_income = original_income

	determined_income = determined_income & ""
	ButtonPressed = income_calc_btn
end function

function app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, ACCOUNTS_ARRAY)
	return_btn = 5001
	enter_btn = 5002
	add_another_btn = 5003
	remove_one_btn = 5004
	assets_review_completed = True

	original_assets = determined_assets
	determined_assets = 0
	If cash_amount_yn <> "Yes" OR bank_account_yn <> "Yes" Then
		Do
			prvt_err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 271, 135, "Determination of Assets in Month of Application"
			  Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
			  GroupBox 10, 25, 255, 40, "Cash"
			  Text 25, 45, 155, 10, "Does the household have any Cash Savings?"
			  DropListBox 180, 40, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", cash_amount_yn
			  GroupBox 10, 70, 255, 40, "Accounts"
			  Text 20, 90, 190, 10, "Does anyone in the household have any Bank Accounts?"
			  DropListBox 210, 85, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", bank_account_yn
			  ButtonGroup ButtonPressed
			    PushButton 215, 115, 50, 15, "Enter", enter_btn
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then
				assets_review_completed = False
				Exit Do
			End If

			If cash_amount_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has CASH."
			If bank_account_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has A BANK ACCOUNT."

			If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
		Loop until prvt_err_msg = ""
	End If

	If assets_review_completed = True Then
		Do
			prvt_err_msg = ""
			cash_amount = cash_amount & ""

			If cash_amount_yn = "No" Then cash_grp_len = 30
			If cash_amount_yn = "Yes" Then cash_grp_len = 50
			If bank_account_yn = "No" Then acct_grp_len = 30
			If bank_account_yn = "Yes" Then acct_grp_len = 60 + (UBound(ACCOUNTS_ARRAY, 2) + 1) * 20
			dlg_len = 55 + cash_grp_len + acct_grp_len

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 351, dlg_len, "Determination of Assets in Month of Application"
			  Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
			  GroupBox 10, 25, 220, cash_grp_len, "Cash"
			  If cash_amount_yn = "Yes" Then
				  Text 20, 40, 155, 10, "This household HAS Cash Savings."
				  Text 20, 55, 150, 10, "How much in total does the household have?"
				  EditBox 175, 50, 45, 15, cash_amount
				  y_pos = 80
			  Else
				  Text 20, 40, 155, 10, "This household does NOT have Cash."
				  y_pos = 60
			  End If
			  GroupBox 10, y_pos, 335, acct_grp_len, "Accounts"
			  y_pos = y_pos + 15
			  If bank_account_yn = "Yes" Then
				  Text 20, y_pos, 190, 10, "This household HAS Bank Accounts."
				  y_pos = y_pos + 15
				  Text 20, y_pos, 50, 10, "Account Type"
				  Text 90, y_pos, 70, 10, "Owner of Account"
				  Text 180, y_pos, 45, 10, "Bank Name"
				  Text 285, y_pos, 35, 10, "Amount"
				  y_pos = y_pos + 15

				  For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
					  ACCOUNTS_ARRAY(account_amount_const, the_acct) = ACCOUNTS_ARRAY(account_amount_const, the_acct) & ""
					  DropListBox 20, y_pos, 60, 45, "Select One..."+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Other", ACCOUNTS_ARRAY(account_type_const, the_acct)
					  EditBox 90, y_pos, 85, 15, ACCOUNTS_ARRAY(account_owner_const, the_acct)
					  EditBox 180, y_pos, 100, 15, ACCOUNTS_ARRAY(bank_name_const, the_acct)
					  EditBox 285, y_pos, 50, 15, ACCOUNTS_ARRAY(account_amount_const, the_acct)
					  y_pos = y_pos + 20
				  Next
			  Else
			  	  Text 20, y_pos, 155, 10, "This household does NOT have Bank Accounts."
			  End If
			  ButtonGroup ButtonPressed
			    If bank_account_yn = "Yes" Then PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_btn
			    If bank_account_yn = "Yes" Then PushButton 275, y_pos, 60, 10, "REMOVE ONE", remove_one_btn
				PushButton 295, dlg_len - 20, 50, 15, "Return", return_btn
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then
				assets_review_completed = False
				Exit Do
			End If
			If ButtonPressed = -1 Then ButtonPressed = return_btn

			last_acct_item = UBound(ACCOUNTS_ARRAY, 2)
			If ButtonPressed = add_another_btn Then
				last_acct_item = last_acct_item + 1
				ReDim Preserve ACCOUNTS_ARRAY(account_notes_const, last_acct_item)
			End If
			If ButtonPressed = remove_one_btn Then
				last_acct_item = last_acct_item - 1
				ReDim Preserve ACCOUNTS_ARRAY(account_notes_const, last_acct_item)
			End If

			cash_amount = trim(cash_amount)
			If cash_amount <> "" And IsNumeric(cash_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the Cash Amount as a number."

			For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
				ACCOUNTS_ARRAY(account_amount_const, the_acct) = trim(ACCOUNTS_ARRAY(account_amount_const, the_acct))
				If ACCOUNTS_ARRAY(account_amount_const, the_acct) <> "" And IsNumeric(ACCOUNTS_ARRAY(account_amount_const, the_acct)) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the Bank Account amounts as a member."
				If ACCOUNTS_ARRAY(account_type_const, the_acct)	= "Select One..." Then prvt_err_msg = prvt_err_msg & vbCr & "* Select the Bank Account type."
			Next
			If prvt_err_msg <> "" AND ButtonPressed = return_btn Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
		Loop Until ButtonPressed = return_btn AND prvt_err_msg = ""

		If cash_amount = "" Then cash_amount = 0
		cash_amount = cash_amount * 1
		For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
			If ACCOUNTS_ARRAY(account_amount_const, the_acct) = "" Then ACCOUNTS_ARRAY(account_amount_const, the_acct) = 0
			ACCOUNTS_ARRAY(account_amount_const, the_acct) = ACCOUNTS_ARRAY(account_amount_const, the_acct) * 1
			determined_assets = determined_assets + ACCOUNTS_ARRAY(account_amount_const, the_acct)
		Next
		determined_assets = determined_assets + cash_amount
	End If
	If assets_review_completed = False Then determined_assets =  original_assets

	determined_assets = determined_assets & ""
	ButtonPressed = asset_calc_btn
end function

function app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)
	return_btn = 5001

	shel_review_completed = True
	rent_amount = rent_amount & ""
	lot_rent_amount = lot_rent_amount & ""
	mortgage_amount = mortgage_amount & ""
	insurance_amount = insurance_amount & ""
	tax_amount = tax_amount & ""
	room_amount = room_amount & ""
	garage_amount = garage_amount & ""
	subsidy_amount = subsidy_amount & ""

	original_shel = determined_shel
	determined_shel = 0
	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 196, 140, "Determination of Housing Cost in Month of Application"
		  EditBox 45, 35, 50, 15, rent_amount
		  EditBox 45, 55, 50, 15, lot_rent_amount
		  EditBox 45, 75, 50, 15, mortgage_amount
		  EditBox 45, 95, 50, 15, insurance_amount
		  EditBox 140, 35, 50, 15, tax_amount
		  EditBox 140, 55, 50, 15, room_amount
		  EditBox 140, 75, 50, 15, garage_amount
		  EditBox 140, 95, 50, 15, subsidy_amount
		  ButtonGroup ButtonPressed
		    PushButton 140, 120, 50, 15, "Return", return_btn
		  Text 10, 15, 165, 10, "Enter the total Shelter Expense for the Houshold."
		  Text 25, 40, 20, 10, "Rent:"
		  Text 10, 60, 35, 10, " Lot Rent:"
		  Text 10, 80, 35, 10, "Mortgage:"
		  Text 10, 100, 35, 10, "Insurance:"
		  Text 115, 40, 25, 10, "Taxes:"
		  Text 115, 60, 25, 10, "Room:"
		  Text 110, 80, 30, 10, "Garage:"
		  Text 105, 100, 35, 10, "  Subsidy:"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then
			shel_review_completed = False
			Exit Do
		End If

		rent_amount = trim(rent_amount)
		lot_rent_amount = trim(lot_rent_amount)
		mortgage_amount = trim(mortgage_amount)
		insurance_amount = trim(insurance_amount)
		tax_amount = trim(tax_amount)
		room_amount = trim(room_amount)
		garage_amount = trim(garage_amount)
		subsidy_amount = trim(subsidy_amount)

		If rent_amount <> "" AND IsNumeric(rent_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the RENT amount as a number."
		If lot_rent_amount <> "" AND IsNumeric(lot_rent_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the LOT RENT amount as a number."
		If mortgage_amount <> "" AND IsNumeric(mortgage_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the MORTGAGE amount as a number."
		If insurance_amount <> "" AND IsNumeric(insurance_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the INSURANCE amount as a number."
		If tax_amount <> "" AND IsNumeric(tax_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the TAXES amount as a number."
		If room_amount <> "" AND IsNumeric(room_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the ROOM amount as a number."
		If garage_amount <> "" AND IsNumeric(garage_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the GARAGE amount as a number."
		If subsidy_amount <> "" AND IsNumeric(subsidy_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the SUBSIDY amount as a number."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If IsNumeric(rent_amount) = True Then determined_shel = determined_shel + rent_amount
	If IsNumeric(lot_rent_amount) = True Then determined_shel = determined_shel + lot_rent_amount
	If IsNumeric(mortgage_amount) = True Then determined_shel = determined_shel + mortgage_amount
	If IsNumeric(insurance_amount) = True Then determined_shel = determined_shel + insurance_amount
	If IsNumeric(tax_amount) = True Then determined_shel = determined_shel + tax_amount
	If IsNumeric(room_amount) = True Then determined_shel = determined_shel + room_amount
	If IsNumeric(garage_amount) = True Then determined_shel = determined_shel + garage_amount
	' If IsNumeric(subsidy_amount) = True Then determined_shel = determined_shel + subsidy_amount

	If shel_review_completed = False Then determined_shel = original_shel

	determined_shel = determined_shel & ""
	ButtonPressed = housing_calc_btn
end function

function app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
	calculate_btn = 5000
	return_btn = 5001
	determined_utilities = 0
	If heat_expense = True then heat_checkbox = checked
	If ac_expense = True then ac_checkbox = checked
	If electric_expense = True then electric_checkbox = checked
	If phone_expense = True then phone_checkbox = checked
	If none_expense = True then none_checkbox = checked

	Do
		current_utilities = all_utilities

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 246, 175, "Determination of Utilities in Month of Application"
		  CheckBox 30, 45, 50, 10, "Heat", heat_checkbox
		  CheckBox 30, 60, 65, 10, "Air Conditioning", ac_checkbox
		  CheckBox 30, 75, 50, 10, "Electric", electric_checkbox
		  CheckBox 30, 90, 50, 10, "Phone", phone_checkbox
		  CheckBox 30, 105, 50, 10, "NONE", none_checkbox
		  ButtonGroup ButtonPressed
		    PushButton 170, 105, 65, 15, "Calculate", calculate_btn
		    PushButton 170, 155, 65, 15, "Return", return_btn
		  Text 10, 10, 235, 10, "Check the boxes for each utility the household is responsible to pay:"
		  GroupBox 15, 30, 225, 95, "Utilities"
		  Text 150, 45, 50, 10, "$ " & determined_utilities
		  Text 150, 60, 35, 35, all_utilities
		  Text 15, 135, 225, 20, "Remember, this expense could be shared, they are still considered responsible to pay and we count the WHOLE standard."
		EndDialog

		dialog Dialog1

		some_vs_none_discrepancy = False
		If (heat_checkbox = checked OR ac_checkbox = checked OR electric_checkbox = checked OR phone_checkbox = checked) AND none_checkbox = checked Then some_vs_none_discrepancy = True
		If some_vs_none_discrepancy = True Then MsgBox "Attention:" & vbCr & vbCr & "You have selected NONE and selected at least one other utility expense. If it is NONE, then no other utilities should be checked."

		all_utilities = ""
		If heat_checkbox = checked Then all_utilities = all_utilities & ", Heat"
		If ac_checkbox = checked Then all_utilities = all_utilities & ", AC"
		If electric_checkbox = checked Then all_utilities = all_utilities & ", Electric"
		If phone_checkbox = checked Then all_utilities = all_utilities & ", Phone"
		If none_checkbox = checked Then all_utilities = all_utilities & ", None"
		If left(all_utilities, 2) = ", " Then all_utilities = right(all_utilities, len(all_utilities) - 2)

		If all_utilities = current_utilities AND ButtonPressed = -1 Then ButtonPressed = return_btn

		determined_utilities = 0
		If heat_checkbox = checked OR ac_checkbox = checked Then
			determined_utilities = determined_utilities + heat_AC_amt
		Else
			If electric_checkbox = checked Then determined_utilities = determined_utilities + electric_amt
			If phone_checkbox = checked Then determined_utilities = determined_utilities + phone_amt
		End If

	Loop Until ButtonPressed = return_btn And some_vs_none_discrepancy = False

	heat_expense = False
	ac_expense = False
	electric_expense = False
	phone_expense = False
	none_expense = False

	If heat_checkbox = checked Then heat_expense = True
	If ac_checkbox = checked Then ac_expense = True
	If electric_checkbox = checked Then electric_expense = True
	If phone_checkbox = checked Then phone_expense = True
	If none_checkbox = checked Then none_expense = True

	ButtonPressed = utility_calc_btn
end function

Function determine_actions(case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, is_elig_XFS, snap_denial_date, approval_date, date_of_application, do_we_have_applicant_id, action_due_to_out_of_state_benefits, mn_elig_begin_date, other_snap_state, case_has_previously_postponed_verifs_that_prevent_exp_snap, delay_action_due_to_faci, deny_snap_due_to_faci)

	case_assesment_text = ""
	next_steps_one = ""
	next_steps_two = ""
	next_steps_three = ""
	next_steps_four = ""
	If IsDate(snap_denial_date) = True Then
		case_assesment_text = "DENIAL has been determined - Case does not meet 'All Other Eligibility Criteria'."
		next_steps_one = "Complete the DENIAL by updating MAXIS and enter a full, detaild DENIAL CASE/NOTE. Complete the full processing before moving on to your next task."

		If action_due_to_out_of_state_benefits = "DENY" Then
			add_msg = "Update MEMI with out of state benefit information to generate accurate DENIAL Results. Add a WCOM to the denial advising resident to reapply within 30 days of the benefits ending in the other state."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If
		End If
		If deny_snap_due_to_faci = True Then
			add_msg = "Ensure FACI is coded correctly for accurate DENIAL. Add a WCOM to the denials advising resident to rapply when release from the facility is within 30 days."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If
		End If

		add_msg = "Process this denial quickly as a PENDING SNAP case will continue to be assigned until acted on, once the determination is done and action can be taken, we do not want to reassign this case."
		If next_steps_two = "" then
			next_steps_two = add_msg
		ElseIf next_steps_three = "" Then
			next_steps_three = add_msg
		ElseIf next_steps_four = "" Then
			next_steps_four = add_msg
		End If

		add_msg = "Denials can be coded in REPT/PND2 if they are for a resident 'Withdraw' of their request. Otherise, since the interview should be done at this point, denials should be processed in STAT."
		If next_steps_three = "" Then
			next_steps_three = add_msg
		ElseIf next_steps_four = "" Then
			next_steps_four = add_msg
		End If

		add_msg = "It is best practice to add detail to the Denial WCOM for clarity for the resident."
		If next_steps_four = "" Then next_steps_four = add_msg
	ElseIf is_elig_XFS = True Then
		If IsDate(approval_date) = True Then
			case_assesment_text = "Case IS EXPEDITED and ready to approve"
			next_steps_one = "Approve SNAP Expedited package of " & expedited_package & " before moving on to your next task. Update MAXIS STAT panels to generate EXPEDITED SNAP Eligibility Results and APPROVE."

			If action_due_to_out_of_state_benefits = "APPROVE" AND mn_elig_begin_date <> date_of_application Then
				If DateDiff("d", date, mn_elig_begin_date) > 0 Then
					add_msg = "After approval, send a BENE request in SIR to have benefits issued on " & mn_elig_begin_date & " instead of the regular issuance day."
					If next_steps_two = "" then
						next_steps_two = add_msg
					ElseIf next_steps_three = "" Then
						next_steps_three = add_msg
					ElseIf next_steps_four = "" Then
						next_steps_four = add_msg
					End If
				End If
			End If

			add_msg = "Remember, EXPEDITED is based on income, assets, and shelter/utility expenses only. Even having a delay reason does not mean the case is not still EXPEDITED."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "We attempt to approve expedited within 24 hours of the date of application, or as close to that time as possible. It is crucial we complete the approval at the time we determine the case to be EXPEDITED."
			If next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "EBT Card information can be found below, but often requires contact with the resident, remember REI issuances can prevent residents from receiving their card."
			If next_steps_four = "" Then next_steps_four = add_msg
		Else
			case_assesment_text = "Case IS EXPEDITED but approval must be delayed."
			next_steps_one = "We must strive to approve this case for the EXPEDITED package of " & expedited_package & " as soon as possible. Make every effort to complete the requirements of this delay and approve the case"

			If do_we_have_applicant_id = False Then
				add_msg = "Double check the case file for ANY document that can be used as an identity document.Advise resident to get us ANY form of ID they can, MNbenefits or the virtual dropbox may be quickest way to receive this document."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If
			If action_due_to_out_of_state_benefits = "FOLLOW UP" Then
				If other_snap_state <> "" Then add_msg = "Contact " & other_snap_state & " as soon as possible to determine the end date of of SNAP in " & other_snap_state & "."
				If other_snap_state = "" Then add_msg = "Contact the other state as soon as possible to determine the end date of of SNAP in that state."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If

			If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then
				add_msg = "This case needs regular review to be able to approve SNAP as soon as, the current verifications come in OR the previous verifications come in. Assist the resident in getting any verifications that we can."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If

			If delay_action_due_to_faci = True Then
				add_msg = "Advise resident and the facility to contact us as soon as possible to be able to approve SNAP once the resident leaves the facility."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If

			add_msg = "Delays in processing Expedited should be few and far between, we must make every reasonable effort to get these cases approved as quickly as possible."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "Check in with Knowledge Now about this case, as delays cause negative impact on our timeliness reports."
			If next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "Remember, EXPEDITED is based on income, assets, and shelter/utility expenses only. Even having a delay reason does not mean the case is not still EXPEDITED."
			If next_steps_four = "" Then next_steps_four = add_msg
		End If
	ElseIf is_elig_XFS = False Then
		case_assesment_text = "Case is NOT EXPEDITED, approval decision should follow standard SNAP Policy."
		next_steps_one = "If there are mandatory verifications, request them immediately. If all verifications have been received, process the case right away."
		next_steps_two = ""
		next_steps_three = ""
		next_steps_four = ""
	End If
end function

function determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)
	determined_income = trim(determined_income)
	If determined_income = "" Then determined_income = 0
	determined_income = determined_income * 1

	determined_assets = trim(determined_assets)
	If determined_assets = "" Then determined_assets = 0
	determined_assets = determined_assets * 1

	determined_shel = trim(determined_shel)
	If determined_shel = "" Then determined_shel = 0
	determined_shel = determined_shel * 1

	determined_utilities = trim(determined_utilities)
	If determined_utilities = "" Then determined_utilities = 0
	determined_utilities = determined_utilities * 1

	calculated_resources = determined_income + determined_assets
	calculated_expenses = determined_shel + determined_utilities

	calculated_low_income_asset_test = False
	calculated_resources_less_than_expenses_test = False
	is_elig_XFS = False

	If determined_income < 150 AND determined_assets <= 100 Then calculated_low_income_asset_test = True
	If calculated_resources < calculated_expenses Then calculated_resources_less_than_expenses_test = True

	If calculated_low_income_asset_test = True OR calculated_resources_less_than_expenses_test = True Then is_elig_XFS = True

	determined_income = determined_income & ""
	determined_assets = determined_assets & ""
	determined_shel = determined_shel & ""
	determined_utilities = determined_utilities & ""
end function

function snap_in_another_state_detail(date_of_application, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)
	original_snap_denial_date = snap_denial_date
	original_snap_denial_reason = snap_denial_explain
	calculation_done = False
	other_state_benefits_openended = False
	action_due_to_out_of_state_benefits = ""
	' other_snap_state = "MN - Minnesota"
	day_30_from_application = DateAdd("d", 30, date_of_application)
	calculate_btn = 5000
	return_btn = 5001

	Do
		Do
			prvt_err_msg = ""

			Dialog1 = ""
			If calculation_done = False Then BeginDialog Dialog1, 0, 0, 381, 190, "Case Received SNAP in Another State"
			If calculation_done = True Then BeginDialog Dialog1, 0, 0, 381, 295, "Case Received SNAP in Another State"
			  DropListBox 255, 55, 110, 45, "Select One..."+chr(9)+state_list, other_snap_state
			  EditBox 255, 75, 60, 15, other_state_reported_benefit_end_date
			  CheckBox 40, 95, 320, 10, "Check here if resident reports the benefits are NOT ended or it is UKNOWN if they are ended.", other_state_benefits_not_ended_checkbox
			  DropListBox 255, 110, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", other_state_contact_yn
			  EditBox 255, 130, 60, 15, other_state_verified_benefit_end_date
			  ButtonGroup ButtonPressed
			    PushButton 325, 170, 50, 15, "Calculate", calculate_btn
			  Text 10, 10, 365, 10, "If a Household has received SNAP in another state, we may still be able to issue Expedited SNAP in Minnesota. "
			  Text 10, 25, 320, 10, "Complete the following information to get guidance on handling cases with SNAP in another State:"
			  GroupBox 10, 45, 365, 120, "Other State Benefits"
			  Text 20, 60, 235, 10, "What State is the Household / Resident receiving SNAP benefits from?"
			  Text 40, 80, 215, 10, "When is the resident REPORTING benefits ending in this state?"
			  Text 20, 115, 230, 10, "Have you called the other state to confirm / discover the SNAP status?"
			  Text 20, 135, 230, 10, "What end date has been confirmed / verified for the other state SNAP?"

			  If calculation_done = True Then
				  GroupBox 10, 190, 365, 80, "Resolution"
				  If action_due_to_out_of_state_benefits = "DENY" Then
					  Text 20, 205, 205, 20, "SNAP should be denied as the other state end date is AFTER the 30 day processing period of the application in MN."
					  Text 245, 205, 120, 10, "Date of Application: " & date_of_application
					  If IsDate(other_state_verified_benefit_end_date) = True Then
					  	Text 255, 215, 110, 10, "End Of Benefits: " & other_state_verified_benefit_end_date
					  ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
					  	Text 255, 215, 110, 10, "End Of Benefits: " & other_state_reported_benefit_end_date
					  End If
					  ' Text 30, 230, 120, 10, "SNAP Denial Date: " & snap_denial_date
					  ' Text 30, 240, 335, 30, "Denial Reason: " & snap_denial_explain
				  ElseIf action_due_to_out_of_state_benefits = "APPROVE" Then
					  Text 20, 205, 205, 20, "SNAP should be APPROVED "
					  Text 245, 205, 120, 10, "Date of Application: " & date_of_application
					  Text 25, 215, 175, 10, "Eligibility can start in MN as of " & mn_elig_begin_date
					  If other_state_contact_yn <> "Yes" Then
					  	Text 20, 230, 340, 10, "Verification of out of state eligibility end can be postponed "
						Text 20, 240, 340, 10, "We should make reasonable efforts to obtain verification so, "
						Text 20, 250, 340, 10, "it is best to attempt a call to the other state right away for verification."
					  End If
				  ElseIf action_due_to_out_of_state_benefits = "FOLLOW UP" Then
					  Text 20, 205, 205, 20, "You must connect with the other state to determine when the benefits have ended or IF the benefits will end."
				  End If
				  ButtonGroup ButtonPressed
				    PushButton 325, 275, 50, 15, "Return", return_btn
			  End If
			EndDialog

			dialog Dialog1

			If ButtonPressed = 0 Then Exit Do

			If IsDate(other_state_reported_benefit_end_date) = False AND other_state_benefits_not_ended_checkbox = unchecked Then prvt_err_msg = prvt_err_msg & vbCr & "* We cannot complete the calculation if a reported end date has not been entered."
			If IsDate(other_state_reported_benefit_end_date) = True AND other_state_benefits_not_ended_checkbox = checked Then prvt_err_msg = prvt_err_msg & vbCr & "* You have entered an end date AND indicated the benefits have not ended by checking the box. Please select only one."

			If IsDate(other_state_reported_benefit_end_date) = True Then
				If DatePart("d", DateAdd("d", 1, other_state_reported_benefit_end_date)) <> 1 Then prvt_err_msg = prvt_err_msg & vbCr & "* SNAP Eligiblity end dates should be the last day of the month that the household received SNAP benefits for. Update the date to be the LAST day of the last month of eligiblity in the other state for the REPORTED end date."
			End If
			If IsDate(other_state_verified_benefit_end_date) = True Then
				If DatePart("d", DateAdd("d", 1, other_state_verified_benefit_end_date)) <> 1 Then prvt_err_msg = prvt_err_msg & vbCr & "* SNAP Eligiblity end dates should be the last day of the month that the household received SNAP benefits for. Update the date to be the LAST day of the last month of eligiblity in the other state for the VERIFIED end date."
			End If
			If prvt_err_msg <> "" Then
				MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
				calculation_done = False
			End If

		Loop until prvt_err_msg = ""

		If ButtonPressed = 0 Then
			calculation_done = False
			Exit Do
		End If

		calculation_done = True
		If other_snap_state = "NB - MN Newborn" OR other_snap_state = "MN - Minnesota" OR other_snap_state = "Select One..." OR other_snap_state = "FC - Foreign Country" OR other_snap_state = "UN - Unknown" Then other_snap_state = ""
		If IsDate(other_state_verified_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_verified_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_verified_benefit_end_date)
				' If DateDiff("d", mn_elig_begin_date, date_of_application) > 0 Then
				' 	mn_elig_begin_date = date_of_application
				' 	expedited_package = original_expedited_package
				' Else
				' 	MN_elig_month = DatePart("m", mn_elig_begin_date)
				' 	MN_elig_month = right("0"&MN_elig_month, 2)
				' 	MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
				' 	expedited_package = MN_elig_month & "/" & MN_elig_year
				' End If
			End If
		ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_reported_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_reported_benefit_end_date)
				' If DateDiff("d", mn_elig_begin_date, date_of_application) > 0 Then
				' 	mn_elig_begin_date = date_of_application
				' 	expedited_package = original_expedited_package
				' Else
				' 	MN_elig_month = DatePart("m", mn_elig_begin_date)
				' 	MN_elig_month = right("0"&MN_elig_month, 2)
				' 	MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
				' 	expedited_package = MN_elig_month & "/" & MN_elig_year
				' End If
			End If
		ElseIf other_state_benefits_not_ended_checkbox = checked Then
			action_due_to_out_of_state_benefits = "FOLLOW UP"
			other_state_benefits_openended = True
		End If
		If action_due_to_out_of_state_benefits <> "DENY" Then
			snap_denial_date = original_snap_denial_date
			snap_denial_explain = original_snap_denial_reason
		End If
		If action_due_to_out_of_state_benefits <> "APPROVE" Then expedited_package = original_expedited_package
	Loop until ButtonPressed = return_btn
	If action_due_to_out_of_state_benefits = "APPROVE" Then
		If DateDiff("d", mn_elig_begin_date, date_of_application) > 0 Then
			mn_elig_begin_date = date_of_application
			expedited_package = original_expedited_package
		Else
			MN_elig_month = DatePart("m", mn_elig_begin_date)
			MN_elig_month = right("0"&MN_elig_month, 2)
			MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
			expedited_package = MN_elig_month & "/" & MN_elig_year
		End If
	End If
	If action_due_to_out_of_state_benefits = "DENY" Then
		snap_denial_date = date
		If other_snap_state = "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in other state. Household can reapply once the eligibility in another state is ending within 30 days"
		If other_snap_state <> "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in " & other_snap_state & ". Household can reapply once the eligibility in another state is ending within 30 days"
		If InStr(snap_denial_explain, deny_msg) = 0 Then snap_denial_explain = snap_denial_explain & "; " & deny_msg & "."
	End If
	If action_due_to_out_of_state_benefits <> "DENY" Then
		If other_snap_state = "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in other state. Household can reapply once the eligibility in another state is ending within 30 days"
		If other_snap_state <> "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in " & other_snap_state & ". Household can reapply once the eligibility in another state is ending within 30 days"
		snap_denial_explain = replace(snap_denial_explain, deny_msg, "")
	End If
	snap_denial_date = snap_denial_date & ""
	ButtonPressed = snap_active_in_another_state_btn
end function

function previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
	fn_review_btn = 5005
	return_btn = 5001
	prev_post_verif_assessment_done = True
	case_has_previously_postponed_verifs_that_prevent_exp_snap = False

	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 446, 160, "Case Previously Received EXP SNAP with Postponed Verifications"
		  Text 10, 10, 435, 10, "A case that was approved Expedited SNAP with postponed verifications MAY not be able to have Expedited Approved right away."
		  Text 10, 30, 125, 10, "This does not apply to cases where:"
		  Text 15, 40, 165, 10, "- The Postponed Verification were not mandatory."
		  Text 15, 50, 275, 10, "- The Postponed Verification were provided - even if Eligibility was not approved."
		  Text 15, 60, 385, 10, "- The case met all criteria for Regular SNAP to be issued and was approved for 'Ongoing' SNAP for at least one month."
		  Text 15, 85, 175, 15, "What is the DATE OF APPLICATION for the Expedited Approval that had Postponed Verifications?"
		  EditBox 195, 85, 50, 15, previous_date_of_application
		  Text 275, 110, 115, 10, "Are these verifications mandatory?"
		  DropListBox 400, 105, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", prev_verifs_mandatory_yn
		  Text 15, 110, 175, 10, "List the verifications that were previously postponed:"
		  EditBox 15, 120, 425, 15, prev_verif_list
		  Text 15, 145, 220, 10, "Does the case have Postponed Verifications for THIS Application?"
		  DropListBox 235, 140, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", curr_verifs_postponed_yn
		  ButtonGroup ButtonPressed
		    PushButton 390, 140, 50, 15, "Review", fn_review_btn
		EndDialog

		dialog Dialog1

		If ButtonPressed = 0 Then
			prev_post_verif_assessment_done = False
			Exit Do
		End If

		prev_verif_list = trim(prev_verif_list)
		If IsDate(previous_date_of_application) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the date of application from the last time this case received an Expedited SNAP approval WITH Postponed Verifications."
		If prev_verifs_mandatory_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* You must review the verifications that were previously postponed and enter them here."
		If prev_verif_list = "" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review the verifications that were previously postponed and indicate if any of them were mandatory."
		If curr_verifs_postponed_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Indicate if the CURRENT application has verifications required that would need to be postponed to approve the Expedited SNAP."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If prev_post_verif_assessment_done = True Then
		PREVIOUS_footer_month = DatePart("m", previous_date_of_application)
		PREVIOUS_footer_month = right("0"&PREVIOUS_footer_month, 2)

		PREVIOUS_footer_year = right(DatePart("yyyy", previous_date_of_application), 2)

		If DatePart("d", previous_date_of_application) > 15 Then
			second_month_of_previous_exp_package = DateAdd("m", 1, previous_date_of_application)
			PREVIOUS_footer_month = DatePart("m", second_month_of_previous_exp_package)
			PREVIOUS_footer_month = right("0"&PREVIOUS_footer_month, 2)

			PREVIOUS_footer_year = right(DatePart("yyyy", second_month_of_previous_exp_package), 2)
		End If
		previous_expedited_package = PREVIOUS_footer_month & "/" & PREVIOUS_footer_year

		ask_more_questions = False
		If IsDate(previous_date_of_application) = True AND prev_verifs_mandatory_yn = "Yes" AND curr_verifs_postponed_yn = "Yes" Then ask_more_questions = True
		If ask_more_questions = True Then
			Do
				prvt_err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 436, 110, "Case Previously Received EXP SNAP with Postponed Verifications"
				  Text 10, 10, 435, 10, "A case that was approved Expedited SNAP with postponed verifications MAY not be able to have Expedited Approved right away."
				  Text 10, 30, 125, 10, "This does not apply to cases where:"
				  Text 15, 40, 165, 10, "- The Postponed Verification were not mandatory."
				  Text 15, 50, 275, 10, "- The Postponed Verification were provided - even if Eligibility was not approved."
				  Text 15, 60, 385, 10, "- The case met all criteria for Regular SNAP to be issued and was approved for 'Ongoing' SNAP for at least one month."
				  Text 10, 80, 180, 10, "Did the case get approved for any SNAP after " & previous_expedited_package & "?"
				  DropListBox 195, 75, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", ongoing_snap_approved_yn
				  Text 20, 95, 170, 10, "Check ECF, are the postponed verifications on file?"
				  DropListBox 195, 90, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", prev_post_verifs_recvd_yn
				  ButtonGroup ButtonPressed
				    PushButton 380, 90, 50, 15, "Review", fn_review_btn

				  Text 10, 270, 280, 20, "If a case cannot be approved due to previously not received Postponed Verifications, the case must meet ONE of the following criteria:"
				  Text 15, 295, 210, 10, "- Provide all verifications that were postponed and mandatory."
				  Text 15, 305, 280, 10, "- Meet all criterea to approve SNAP - including receipt of all mandatory verifications."
				  Text 20, 315, 265, 20, "(This means if a case has no verifications to request, we CAN approve Expedited as the case meets all criteria to approve SNAP.)"
				EndDialog

				dialog Dialog1

				If ButtonPressed = 0 Then
					prev_post_verif_assessment_done = False
					Exit Do
				End If

				If ongoing_snap_approved_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review MAXIS and determine if SNAP was approved after the last month of the expedited package (" & previous_expedited_package & "). If it was, the case met all requirements to gain SNAP eligibility."
				If prev_post_verifs_recvd_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review the ECF case file and see if the mandatory postponed verifications were ever received, even if SNAP was not approved."

				If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
			Loop until prvt_err_msg = ""
		End If
	End If

	If prev_post_verif_assessment_done = True Then
		If ask_more_questions = False OR ongoing_snap_approved_yn = "Yes" OR prev_post_verifs_recvd_yn = "Yes" Then
			Dialog1 = ""
			y_pos = 85

			BeginDialog Dialog1, 0, 0, 436, 120, "Case Previously Received EXP SNAP with Postponed Verifications"
			  GroupBox 10, 10, 415, 55, "EXPEDITED CAN BE APPROVED"
			  Text 25, 25, 100, 10, "Based on this case situation"
			  Text 30, 35, 325, 10, "This case CAN be approved for Expedited without a delay due to Previous Postponed Verifications."
			  Text 35, 45, 285, 10, "(There may be another reason for delay, complete the rest of the review to determine.)"
			  Text 15, 75, 45, 10, "Explanation:"
			  If prev_verifs_mandatory_yn = "No" Then
				  Text 15, y_pos, 350, 10, "The previously postponed verifications were not mandatory, so case met all SNAP eligibility criteria."
				  y_pos = y_pos + 10
			  End If
			  If curr_verifs_postponed_yn = "No" Then
				  Text 15, y_pos, 350, 10, "There are no verifications that are required and being postponed now, so case meets all SNAP eligibility criteria."
				  y_pos = y_pos + 10
			  End If
			  If ongoing_snap_approved_yn = "Yes" Then
				  Text 15, y_pos, 350, 10, "Case was approved regular SNAP after the expedited package time, so case met all SNAP eligibility criteria."
				  y_pos = y_pos + 10
			  End If
			  If prev_post_verifs_recvd_yn = "Yes" Then
				  Text 50, y_pos, 350, 10, "The postponed verifications have been received, which meets the requirement to receive another posponed verification approval package."
				  y_pos = y_pos + 10
			  End If
			  ButtonGroup ButtonPressed
			    PushButton 380, 100, 50, 15, "Update", update_btn
			EndDialog

			dialog Dialog1

		End If

		If ask_more_questions = True AND ongoing_snap_approved_yn = "No" AND prev_post_verifs_recvd_yn = "No" Then
			case_has_previously_postponed_verifs_that_prevent_exp_snap = True

			BeginDialog Dialog1, 0, 0, 291, 145, "Case Previously Received EXP SNAP with Postponed Verifications"
			  GroupBox 5, 5, 280, 60, "EXPEDITED APPROVAL MUST BE DELAYED"
			  Text 20, 20, 100, 10, "Based on this case situation"
			  Text 25, 30, 195, 10, "This case CANNOT be approved for Expedited at this time."
			  Text 30, 40, 235, 20, "The case would require postponing verifications when we already have allowed for postponed verifications that have not been received."
			  Text 10, 70, 275, 20, "If a case cannot be approved due to previously not received Postponed Verifications, the case must meet ONE of the following criteria:"
			  Text 15, 95, 210, 10, "- Provide all verifications that were postponed and mandatory."
			  Text 15, 105, 280, 10, "- Meet all criterea to approve SNAP - including receipt of all mandatory verifications."
			  Text 20, 115, 265, 20, "(This means if a case has no verifications to request, we CAN approve Expedited as the case meets all criteria to approve SNAP.)"
			  ButtonGroup ButtonPressed
			    PushButton 235, 125, 50, 15, "Update", update_btn
			EndDialog

			dialog Dialog1
		End If
	End If
	If prev_post_verif_assessment_done = False Then
		case_has_previously_postponed_verifs_that_prevent_exp_snap = False
		Explain_not_completed_msg = Msgbox("All of the details around postponed verifications have not been entered to be able to determine if there should be a delay due to previously postponed verifications." & vbCr & vbCr & "If you have details to record and you wish to complete the assesment, press the button for this functionality again and the script will restart the questions.", vbOK, "Escape Pressed - Details not Completed")
	End If
	delay_msg = "Approval cannot be completed as case has postponed verifications when postpone verifications were previously allowed and not provided, nor has the case meet 'ongoing SNAP' eligibility"
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = False Then delay_explanation = replace(delay_explanation, delay_msg, "")
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then
		If InStr(delay_explanation, delay_msg) = 0 Then delay_explanation = delay_explanation & "; " & delay_msg & "."
	End If

	ButtonPressed = case_previously_had_postponed_verifs_btn
end function

function household_in_a_facility_detail(delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, delay_explanation, snap_denial_explain, snap_denial_date, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
	return_btn = 5001
	delay_action_due_to_faci = False
	deny_snap_due_to_faci = False
	faci_review_completed = True

	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 266, 200, "Case Previously Received EXP SNAP with Postponed Verifications"
		  EditBox 70, 40, 180, 15, facility_name
		  DropListBox 210, 60, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", snap_inelig_faci_yn
		  EditBox 110, 100, 50, 15, faci_entry_date
		  EditBox 110, 120, 50, 15, faci_release_date
		  CheckBox 110, 140, 150, 10, "Check here if the release date is unknown.", release_date_unknown_checkbox
		  DropListBox 210, 155, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", release_within_30_days_yn
		  ButtonGroup ButtonPressed
		    PushButton 215, 180, 45, 15, "Return", return_btn
		  Text 10, 10, 90, 10, "Resident is in a Facility"
		  GroupBox 10, 25, 250, 55, "Facility Information"
		  Text 20, 45, 50, 10, "Facility Name"
		  Text 95, 65, 115, 10, "Is this a 'SNAP Ineligible' facility?"
		  GroupBox 10, 85, 250, 90, "Resident Stay Information"
		  Text 20, 105, 85, 10, "Date of Entry into Facility:"
		  Text 30, 125, 75, 10, "Date of Exit / Release:"
		  Text 165, 125, 45, 10, "(or expected)"
		  Text 20, 160, 185, 10, "Does the resident expect to be released by " & day_30_from_application & "?"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then
			faci_review_completed = False
			Exit Do
		End If

		facility_name = trim(facility_name)
		If facility_name = "" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the name of the facility."
		If snap_inelig_faci_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Select if this is a SNAP Ineligible Facility."
		If IsDate(faci_release_date) = False AND release_date_unknown_checkbox = unchecked Then prvt_err_msg = prvt_err_msg & vbCr & "* Either enter a release date (expected release date) or indicate that the release date is unknown."
		If IsDate(faci_release_date) = True AND release_date_unknown_checkbox = checked Then prvt_err_msg = prvt_err_msg & vbCr & "* You have entered a release date AND indicated the release date is unknown."
		If release_date_unknown_checkbox = checked AND release_within_30_days_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Since the expected release date is unknown, indicate if this release is expected to be prior to do the end of the 30 day processing period."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If faci_review_completed = True Then
		If snap_inelig_faci_yn = "Yes" Then
			If IsDate(faci_release_date) = True Then
				If DateDiff("d", date, faci_release_date) > 0 AND DateDiff("d", faci_release_date, day_30_from_application) >= 0 Then delay_action_due_to_faci = True
				If DateDiff("d", date, faci_release_date) > 0 AND DateDiff("d", faci_release_date, day_30_from_application) < 0 Then deny_snap_due_to_faci = True
			ElseIf release_date_unknown_checkbox = checked Then
				If release_within_30_days_yn = "Yes" Then delay_action_due_to_faci = True
				If release_within_30_days_yn = "No" Then deny_snap_due_to_faci = True
 			End If
		End If
	End If

	delay_msg = "Approval cannot be completed as resident is still in an Ineligible SNAP Facility"
	If delay_action_due_to_faci = False Then delay_explanation = replace(delay_explanation, delay_msg, "")
	If delay_action_due_to_faci = True Then
		If InStr(delay_explanation, delay_msg) = 0 Then delay_explanation = delay_explanation & "; " & delay_msg & "."
	End If

	deny_msg = "SNAP to be denied as resident is in an Ineligible SNAP Facility and is not expected to be released within 30 days of the Date of Application"
	If deny_snap_due_to_faci = False Then
		If InStr(snap_denial_explain, deny_msg) = 0 Then snap_denial_date = ""
		snap_denial_explain = replace(snap_denial_explain, deny_msg, "")
	End If
	If deny_snap_due_to_faci = True Then
		If InStr(snap_denial_explain, deny_msg) = 0 Then snap_denial_explain = snap_denial_explain & "; " & deny_msg & "."
		snap_denial_date = date
		snap_denial_date = snap_denial_date & ""
	End If

	If faci_review_completed = True Then
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 216, 130, "Case Previously Received EXP SNAP with Postponed Verifications"
		  Text 10, 10, 90, 10, "Resident is in a Facility"
		  ButtonGroup ButtonPressed
		    PushButton 165, 110, 45, 15, "Return", return_btn
		  Text 15, 25, 140, 20, "The resident's stay in the Facility impacts the SNAP Expedited Processing by:"
		  If delay_action_due_to_faci = True Then Text 20, 55, 195, 10, "Delaying the Approval of Expedited until the Release Date"
		  If deny_snap_due_to_faci = True Then Text 20, 55, 190, 20, "The SNAP case should be DENIED as the resident will not be released within 30 days."
		  If delay_action_due_to_faci = False AND deny_snap_due_to_faci = False Then Text 20, 55, 195, 10, "No change to the Expedited processing because:"
		  y_pos = 65
		  If snap_inelig_faci_yn = "No" Then
			  Text 30, y_pos, 180, 10, "The Facility is not a SNAP Ineligible Facility."
			  y_pos = y_pos + 10
		  End If
		  If IsDate(faci_release_date) = True Then
			  If DateDiff("d", date, faci_release_date) <= 0 Then
			  	Text 30, y_pos, 180, 30, "The release date has already happend. SNAP Eligibility Begin date should be changed to " & faci_release_date & " and processed based on the rest of the case information."
			  End If
		  End If
		EndDialog

		dialog Dialog1
	End If

	ButtonPressed = household_in_a_facility_btn
end function

function send_support_email_to_KN()

	email_subject = "Assistance with Case at SNAP Application - Possible EXP"
	If developer_mode = True Then email_subject = "TESTING RUN - " & email_subject & " - can be deleted"

	email_body = "I am completing a SNAP Expedited Determination." & vbCr & vbCr
	email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr & vbCr
	email_body = email_body & "Amounts currently entered at the Determination:" & vbCr
	email_body = email_body & "Income: $ " & determined_income & vbCr
	email_body = email_body & "Assets: $ " & determined_assets & vbCr
	email_body = email_body & "Housing: $ " & determined_shel & vbCr
	email_body = email_body & "Utilities: $ " & determined_utilities & vbCr & vbCr
	email_body = email_body & "Script Calculations:" & vbCr
	If is_elig_XFS = True Then email_body = email_body & "Case IS EXPEDITED." & vbCr
	If is_elig_XFS = False Then email_body = email_body & "Case is NOT Expedtied." & vbCr
	email_body = email_body & "Unit has less than $150 monthly Gross Income AND $100 or less in assets: " & calculated_low_income_asset_test & vbCr
	email_body = email_body & "Unit's combined resources are less than housing expense: " & calculated_resources_less_than_expenses_test & vbCr & vbCr
	email_body = email_body & "Case Dates/Timelines:" & vbCr
	email_body = email_body & "Date of Application: " & date_of_application & vbCr
	email_body = email_body & "Date of Interview: " & interview_date & vbCr
	email_body = email_body & "Date of Approval: " & approval_date & " (or planned date of approval)" & vbCr
	email_body = email_body & "Processing Delay Explanation: " & delay_explanation & vbCr
	email_body = email_body & "SNAP Denial Date: " & snap_denial_date & vbCr
	email_body = email_body & "Denial Explanation: " & snap_denial_explain & vbCr & vbCr
	email_body = email_body & "Other Information:" & vbCr
	If applicant_id_on_file_yn <> "" AND applicant_id_on_file_yn <> "?" Then email_body = email_body & "Is there an ID on file for the applicant? " & applicant_id_on_file_yn & vbCr
	If applicant_id_through_SOLQ <> "" AND applicant_id_through_SOLQ <> "?" Then email_body = email_body & "Can the Identity of the applicant be cleard through SOLQ/SMI? " & applicant_id_through_SOLQ & vbCr
	If postponed_verifs_yn <> "" AND postponed_verifs_yn <> "?" Then email_body = email_body & "Are there Postponed Verifications for this case? " & postponed_verifs_yn & vbCr
	If trim(list_postponed_verifs) <> "" Then email_body = email_body & "Postponed Verifications: " & list_postponed_verifs & vbCr
	If action_due_to_out_of_state_benefits <> "" Then
		email_body = email_body & "Other SNAP State: " & other_snap_state & vbCr
		email_body = email_body & "Reported End Date: " & other_state_reported_benefit_end_date & vbCr
		If other_state_benefits_openended = True Then email_body = email_body & "End date of SNAP in other state not determined." & vbCr
		email_body = email_body & "Has other State End Date been Confirmed/Verified: " & other_state_contact_yn & vbCr
		email_body = email_body & "Verified End Date: " & other_state_verified_benefit_end_date & vbCr
		email_body = email_body & "Action recommended by script based on information provided: " & action_due_to_out_of_state_benefits & vbCr
	End If
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then email_body = email_body & "It appears this case has postponed verifications from a previous EXP SNAP package that prevent approval of a new Expedited Package." & vbCr & vbCr

	email_body = email_body & "---" & vbCr
	If worker_name <> "" Then email_body = email_body & "Signed, " & vbCr & worker_name

	email_body = "~~This email is generated from within the 'Expedited Determination' Script.~~" & vbCr & vbCr & email_body
	call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", True)
	' call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", False)
	' create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
end function

function view_poli_temp(temp_one, temp_two, temp_three, temp_four)
	call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
	EMWriteScreen "TEMP", 5, 40     'Writes TEMP

	'Writes the panel_title selection
	Call write_value_and_transmit("TABLE", 21, 71)

	If temp_one <> "" Then temp_one = right("00" & temp_one, 2)
	If len(temp_two) = 1 Then temp_two = right("00" & temp_two, 2)
	If len(temp_three) = 1 Then temp_three = right("00" & temp_three, 2)
	If len(temp_four) = 1 Then temp_four = right("00" & temp_four, 2)

	total_code = "TE" & temp_one & "." & temp_two
	If temp_three <> "" Then total_code = total_code & "." & temp_three
	If temp_four <> "" Then total_code = total_code & "." & temp_four

	EMWriteScreen total_code, 3, 21
	transmit

	EMWriteScreen "X", 6, 4
	transmit
end function
'---------------------------------------------------------------------------------------------------------------------------

'DECLARATIONS---------------------------------------------------------------------------------------------------------------

next_btn = 2
finish_btn = 3

amounts_btn 		= 10
determination_btn 	= 20
review_btn 			= 30

income_calc_btn								= 100
asset_calc_btn								= 110
housing_calc_btn							= 120
utility_calc_btn							= 130
snap_active_in_another_state_btn			= 140
case_previously_had_postponed_verifs_btn	= 150
household_in_a_facility_btn					= 160

knowledge_now_support_btn		= 500
te_02_10_01_btn					= 510

hsr_manual_expedited_snap_btn 	= 1000
hsr_snap_applications_btn		= 1100
ryb_exp_identity_btn			= 1200
ryb_exp_timeliness_btn			= 1300
sir_exp_flowchart_btn			= 1400
cm_04_04_btn					= 1500
cm_04_06_btn					= 1600
ht_id_in_solq_btn				= 1700
cm_04_12_btn					= 1800
temp_prog_changes_ebt_card_btn 	= 1900

const account_type_const	= 0
const account_owner_const	= 1
const bank_name_const		= 2
const account_amount_const	= 3
const account_notes_const 	= 4

Dim ACCOUNTS_ARRAY
ReDim ACCOUNTS_ARRAY(account_notes_const, 0)

const jobs_employee_const 	= 0
const jobs_employer_const	= 1
const jobs_wage_const		= 2
const jobs_hours_const		= 3
const jobs_frequency_const 	= 4
const jobs_monthly_pay_const= 5
const jobs_notes_const 		= 6

Dim JOBS_ARRAY
ReDim JOBS_ARRAY(jobs_notes_const, 0)

const busi_owner_const 				= 0
const busi_info_const 				= 1
const busi_monthly_earnings_const	= 2
const busi_annual_earnings_const	= 3
const busi_notes_const 				= 4

Dim BUSI_ARRAY
ReDim BUSI_ARRAY(busi_notes_const, 0)

const unea_owner_const 				= 0
const unea_info_const 				= 1
const unea_monthly_earnings_const	= 2
const unea_weekly_earnings_const	= 3
const unea_notes_const 				= 4

Dim UNEA_ARRAY
ReDim UNEA_ARRAY(unea_notes_const, 0)

'---------------------------------------------------------------------------------------------------------------------------

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
'connecting to MAXIS & searches for the case number
EMConnect ""

Call check_for_MAXIS(false)
call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

Call find_user_name(worker_name)

'dialog to gather the Case Number and such
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 291, 95, "SNAP EXP Determination - Case Information"
  EditBox 85, 5, 60, 15, MAXIS_case_number
  DropListBox 85, 25, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", maxis_updated_yn
  EditBox 85, 55, 200, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 75, 50, 15
    CancelButton 235, 75, 50, 15
  Text 30, 10, 50, 10, "Case Number:"
  Text 20, 30, 65, 10, "MAXIS Updated?"
  Text 85, 40, 200, 10, "(All income, asset, and expense information entered in STAT)"
  Text 10, 60, 70, 10, "Sign your case note:"
EndDialog


Do
	Do
		Dialog Dialog1
		cancel_without_confirmation
		err_msg = ""
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your worker signature"
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If maxis_updated_yn = "?" Then err_msg = err_msg & vbCr & "* Indicate if MAXIS has been updated with the known information about income, assets, and expenses"
		IF err_msg <> "" THEN MsgBox "***** Action Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false				'loops until user passwords back in

exp_screening_note_found = False
snap_elig_results_read = False
do_we_have_applicant_id = False
developer_mode = False

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
End If
If MX_region = "TRAINING" Then developer_mode = True
' If user_ID_for_validation = "CALO001" OR user_ID_for_validation = "ILFE001" OR user_ID_for_validation = "WFS395"Then
' 	stay_in_dev_mode = MsgBox("HELLO BZ Script Writer!" & vbCr & vbCr & "You are running in DEVELOPER MODE! This means your data will not be stored in the report out." &vbCr & vbCr & "Is this what you want?" & vbCr & "Click 'No' to turn developer mode OFF.", vbQuestion + vbYesNo, "Run in DEVELOPER MODE??")
' 	if stay_in_dev_mode = vbNo Then developer_mode = False
' End If
' Call navigate_to_MAXIS_screen("STAT", "PROG")
Do
	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
	If is_this_priv = True Then Call script_end_procedure("This case is PRIVILEGED and cannot be accessed. Request access to the case first and retry the script once you have access to the case.")
	EMReadScreen panel_prog_check, 4, 2, 50
Loop until panel_prog_check = "PROG"
EMReadScreen case_pw, 7, 21, 21

EMReadScreen date_of_application, 8, 10, 33
EMReadScreen interview_date, 8, 10, 55

date_of_application = replace(date_of_application, " ", "/")
interview_date = replace(interview_date, " ", "/")
If interview_date = "__/__/__" Then interview_date = ""


Do
	Do
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 156, 70, "SNAP EXP Determination - Application Information"
		  EditBox 90, 5, 60, 15, date_of_application
		  EditBox 90, 25, 60, 15, interview_date
		  Text 20, 10, 65, 10, "Date of Application:"
		  Text 25, 30, 60, 10, "Date of Interview:"
		  ButtonGroup ButtonPressed
		    OkButton 45, 50, 50, 15
		    CancelButton 100, 50, 50, 15
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation
		err_msg = ""
		If IsDate(date_of_application) = False Then
			err_msg = err_msg & vbCr & "* The date of application needs to be entered as a valid date."
		Else
			If DateDiff("d", date_of_application, date) < 0 Then err_msg = err_msg & vbCr & "* The Application Date cannot be a Future date."
		End If
		If IsDate(interview_date) = False Then
			err_msg = err_msg & vbCr & "* The interview date needs to be entered as a valid date. An Expedited Determination cannot be completed without the interview."
		Else
			If DateDiff("d", interview_date, date) < 0 Then err_msg = err_msg & vbCr & "* The Interview Date cannot be a Future date."
		End If
		If IsDate(date_of_application) = True AND IsDate(interview_date) = True Then
			' MsgBox DateDiff("d", interview_date, date_of_application)
			If DateDiff("d", interview_date, date_of_application) > 0 Then err_msg = err_msg & vbCr & "* The Interview Date Cannot be before the Application Date."
		End If
		IF err_msg <> "" THEN MsgBox "***** Action Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false

day_30_from_application = DateAdd("d", 30, date_of_application)

MAXIS_footer_month = DatePart("m", date_of_application)
MAXIS_footer_month = right("0"&MAXIS_footer_month, 2)

MAXIS_footer_year = right(DatePart("yyyy", date_of_application), 2)

expedited_package = MAXIS_footer_month & "/" & MAXIS_footer_year
If DatePart("d", date_of_application) > 15 Then
	second_month_of_exp_package = DateAdd("m", 1, date_of_application)
	NEXT_footer_month = DatePart("m", second_month_of_exp_package)
	NEXT_footer_month = right("0"&NEXT_footer_month, 2)

	NEXT_footer_year = right(DatePart("yyyy", second_month_of_exp_package), 2)
	expedited_package = expedited_package & " and " & NEXT_footer_month & "/" & NEXT_footer_year
End If
original_expedited_package = expedited_package

Call hest_standards(heat_AC_amt, electric_amt, phone_amt, date_of_application)

'Script is going to find information that was writen in an Expedited Screening case note using scripts
navigate_to_MAXIS_screen "CASE", "NOTE"

row = 1
col = 1
EMSearch "Received", row, col
IF row <> 0 THEN
	For look_for_right_note = 57 to 72
		EMReadScreen xfs_screen_note, 18, row, look_for_right_note
        xfs_screen_note = UCase(xfs_screen_note)
		IF xfs_screen_note = "CLIENT APPEARS EXP" or xfs_screen_note = "CLIENT DOES NOT AP" THEN
			exp_screening_note_found = TRUE	'IF the script found a case note with the NOTES - Expedited Screening format - it can find the information used
			IF look_for_right_note = 57 or look_for_right_note = 65 THEN
				EMReadScreen xfs_screening, 32, row, 42
			ElseIf look_for_right_note = 64 OR look_for_right_note = 72 THEN
				EMReadScreen xfs_screening, 31, row, 49
			End If
			EMWriteScreen "x", row, 3
			transmit
			Exit For
		END If
	Next
END IF

'Script is gathering the income/asset/expense information from the XFS Screening note
IF exp_screening_note_found = TRUE THEN
    EMReadScreen xfs_screening, 40, 4, 36
    xfs_screening = replace(xfs_screening, "~", "")
    xfs_screening = trim(xfs_screening)
	xfs_screening = UCase(xfs_screening)
	xfs_screening_display = xfs_screening & ""
	row = 1
	col = 1
	EMSearch "CAF 1 income", row, col
	EMReadScreen caf_one_income, 8, row, 42
	IF IsNumeric(caf_one_income) = True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_income = abs(caf_one_income)
	Else
		caf_one_income = 0
	End If

	row = 1
	col = 1
	EMSearch "CAF 1 liquid assets", row, col
	EMReadScreen caf_one_assets, 8, row, 42
	If IsNumeric(caf_one_assets)= True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_assets = caf_one_assets * 1
	Else
		caf_one_assets = 0
	End If

	caf_one_resources = caf_one_income + caf_one_assets	'Totaling the amounts for the case note

	row = 1
	col = 1
	EMSearch "CAF 1 rent", row, col
	EMReadScreen caf_one_rent, 8, row, 42
	IF IsNumeric(caf_one_rent) = True Then 		'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_rent = abs(caf_one_rent)
	Else
		caf_one_rent = 0
	End If

	row = 1
	col = 1
	EMSearch "Utilities (AMT", row, col
	EMReadScreen caf_one_utilities, 8, row, 42
	If IsNumeric(caf_one_utilities) = True Then 	'If a worker alters this note, we need to default to a number so that the script does not break
		caf_one_utilities = abs(caf_one_utilities)
	Else
		caf_one_utilities = 0
	End If

	caf_one_expenses = caf_one_rent + caf_one_utilities		'Totaling the amounts for a case note

	'The script not adjusts the format so it looks nice
	caf_one_income = FormatNumber(caf_one_income, 2, -1, 0, -1)
	caf_one_assets = FormatNumber(caf_one_assets, 2, -1, 0, -1)
	caf_one_rent = FormatNumber(caf_one_rent, 2, -1, 0, -1)
	caf_one_utilities = FormatNumber(caf_one_utilities, 2, -1, 0, -1)
	caf_one_resources = FormatNumber(caf_one_resources, 2, -1, 0, -1)
	caf_one_expenses = FormatNumber(caf_one_expenses, 2, -1, 0, -1)
	PF3
End IF

day_before_app = DateAdd("d", -1, date_of_application) 'will set the date one day prior to app date'

note_row = 5            'resetting the variables on the loop
note_date = ""
note_title = ""
appt_date = ""
Do
	EMReadScreen note_date, 8, note_row, 6      'reading the note date
	EMReadScreen note_title, 55, note_row, 25   'reading the note header
	note_title = trim(note_title)

	IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
		appt_notc_sent_on = note_date
	ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
		appt_notc_sent_on = note_date
	ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
		EMReadScreen appt_date, 10, note_row, 63
		appt_date = replace(appt_date, "~", "")
		appt_date = trim(appt_date)
		appt_notc_sent_on = note_date
		appt_date_in_note = appt_date
	END IF
	' MsgBox "ROW - " & note_row & vbCr & "APPT SENT ON - " & appt_notc_sent_on & "APPT FOR - " & appt_date_in_note


	IF note_date = "        " then Exit Do
	note_row = note_row + 1
	IF note_row = 19 THEN
		PF8
		note_row = 5
	END IF
	EMReadScreen next_note_date, 8, note_row, 6
	IF next_note_date = "        " then Exit Do
Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'

determined_utilities = ""
If maxis_updated_yn = "No" Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)

If maxis_updated_yn = "Yes" Then

	Call Navigate_to_MAXIS_screen("STAT", "MEMB")
	EMReadScreen id_ver_code, 2, 9, 68
	If id_ver_code <> "__" AND id_ver_code <> "NO" Then applicant_id_on_file_yn = "Yes"
	If id_ver_code = "__" OR id_ver_code = "NO" Then applicant_id_on_file_yn = "No"

	const panel_type_const	= 0
	const panel_memb_const 	= 1
	const panel_inst_const 	= 2

	Dim PANELS_TO_READ_ARRAY()
	ReDim PANELS_TO_READ_ARRAY(panel_inst_const, 0)

	Call navigate_to_MAXIS_screen("STAT", "PNLP")

	cash_amount_yn = "No"
	bank_account_yn = "No"
	jobs_income_yn = "No"
	busi_income_yn = "No"
	unea_income_yn = "No"

	rent_amount = 0
	lot_rent_amount = 0
	mortgage_amount = 0
	insurance_amount = 0
	tax_amount = 0
	room_amount = 0
	garage_amount = 0
	subsidy_amount = 0

	determined_shel = 0
	cash_amount = 0
	determined_assets = 0
	determined_income = 0

	ReDim Preserve PANELS_TO_READ_ARRAY(panel_inst_const, 0)
	PANELS_TO_READ_ARRAY(panel_type_const, 0) = "MEMI"
	PANELS_TO_READ_ARRAY(panel_memb_const, 0) = "01"

	income_review_completed = True
	assets_review_completed = True
	shel_review_completed = True

	all_the_panels_looked_at = False
	pnl_row = 3
	array_incrementer = 1
	acct_incrementor = 0
	jobs_incrementor = 0
	unea_incrementor = 0
	busi_incrementor = 0

	instance_counter = 1
	Do
		EMReadScreen the_panel_name, 4, pnl_row, 5
		EMReadScreen the_memb, 2, pnl_row, 10
		' EMReadScreen the_instance, 2, pnl_row, 10

		If the_panel_name = "HEST" Then
			ReDim Preserve PANELS_TO_READ_ARRAY(panel_inst_const, array_incrementer)
			PANELS_TO_READ_ARRAY(panel_type_const, array_incrementer) = the_panel_name

			array_incrementer = array_incrementer + 1
		End If

		' If the_panel_name = "SHEL" OR the_panel_name = "CASH" OR the_panel_name = "MEMI" Then
		If the_panel_name = "SHEL" OR the_panel_name = "CASH" Then
			ReDim Preserve PANELS_TO_READ_ARRAY(panel_inst_const, array_incrementer)
			PANELS_TO_READ_ARRAY(panel_type_const, array_incrementer) = the_panel_name
			PANELS_TO_READ_ARRAY(panel_memb_const, array_incrementer) = the_memb

			array_incrementer = array_incrementer + 1
		End If

		If the_panel_name = "FACI" OR the_panel_name = "ACCT" OR the_panel_name = "JOBS" OR the_panel_name = "BUSI" OR the_panel_name = "UNEA" Then
			If the_panel_name <> last_panel OR the_memb <> last_memb Then instance_counter = 1
			ReDim Preserve PANELS_TO_READ_ARRAY(panel_inst_const, array_incrementer)
			PANELS_TO_READ_ARRAY(panel_type_const, array_incrementer) = the_panel_name
			PANELS_TO_READ_ARRAY(panel_memb_const, array_incrementer) = the_memb
			PANELS_TO_READ_ARRAY(panel_inst_const, array_incrementer) = "0" & instance_counter

			instance_counter = instance_counter + 1
			array_incrementer = array_incrementer + 1
		End If
		last_panel = the_panel_name
		last_memb = the_memb

		pnl_row = pnl_row + 1
		If pnl_row = 20 Then
			transmit
			pnl_row = 3
			EMReadScreen SELF_check, 4, 2, 50
		End If
	Loop Until SELF_check = "SELF"

	Call navigate_to_MAXIS_screen("STAT", "SUMM")
	For each_panel = 0 to UBound(PANELS_TO_READ_ARRAY, 2)
		EMWriteScreen PANELS_TO_READ_ARRAY(panel_type_const, each_panel), 20, 71
		EMWriteScreen PANELS_TO_READ_ARRAY(panel_memb_const, each_panel), 20, 76
		EMWriteScreen PANELS_TO_READ_ARRAY(panel_inst_const, each_panel), 20, 79
		transmit

		EMReadScreen VAR, 8, row, col

		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "HEST" Then
			heat_expense = False
			ac_expense = False
			electric_expense = False
			phone_expense = False

			EMReadScreen heat_ac_yn, 1, 13, 60
			EMReadScreen elec_yn, 1, 14, 60
			EMReadScreen phone_yn, 1, 15, 60

			If heat_ac_yn = "Y" Then heat_expense = True
			If heat_ac_yn = "Y" Then ac_expense = True
			If elec_yn = "Y" Then electric_expense = True
			If phone_yn = "Y" Then phone_expense = True

			all_utilities = ""
			If heat_expense = True Then all_utilities = all_utilities & ", Heat"
			If ac_expense = True Then all_utilities = all_utilities & ", AC"
			If electric_expense = True Then all_utilities = all_utilities & ", Electric"
			If phone_expense = True Then all_utilities = all_utilities & ", Phone"
			If heat_expense = False AND ac_expense = False AND electric_expense = False AND phone_expense = False Then all_utilities = all_utilities & ", None"
			If left(all_utilities, 2) = ", " Then all_utilities = right(all_utilities, len(all_utilities) - 2)


			determined_utilities = 0
			If heat_expense = True OR ac_expense = True Then
				determined_utilities = determined_utilities + heat_AC_amt
			Else
				If electric_expense = True Then determined_utilities = determined_utilities + electric_amt
				If phone_expense = True Then determined_utilities = determined_utilities + phone_amt
			End If
			' Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "MEMI" Then
			EMReadScreen fs_end_other_state, 8, 13, 49
			EMReadScreen mn_entry_date, 8, 15, 49
			EMReadScreen former_state, 2, 15, 78

			fs_end_other_state = replace(fs_end_other_state, " ", "/")
			If fs_end_other_state = "__/__/__" Then fs_end_other_state = ""
			mn_entry_date = replace(mn_entry_date, " ", "/")
			If mn_entry_date = "__/__/__" Then mn_entry_date = ""

			other_state_reported_benefit_end_date = fs_end_other_state
			If former_state <> "__" Then
				state_array = split(state_list, chr(9))
				For each state_item in state_array
					If former_state = left(state_item, 2) Then former_state = state_item
				Next
			End If
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "SHEL" Then
			EMReadscreen panel_rent_amount, 8, 11, 56
			EMReadscreen panel_lot_rent_amount, 8, 12, 56
			EMReadscreen panel_mortgage_amount, 8, 13, 56
			EMReadscreen panel_insurance_amount, 8, 14, 56
			EMReadscreen panel_tax_amount, 8, 15, 56
			EMReadscreen panel_room_amount, 8, 16, 56
			EMReadscreen panel_garage_amount, 8, 17, 56
			EMReadscreen panel_subsidy_amount, 8, 18, 56

			panel_rent_amount = replace(panel_rent_amount, "_", "")
			panel_lot_rent_amount = replace(panel_lot_rent_amount, "_", "")
			panel_mortgage_amount = replace(panel_mortgage_amount, "_", "")
			panel_insurance_amount = replace(panel_insurance_amount, "_", "")
			panel_tax_amount = replace(panel_tax_amount, "_", "")
			panel_room_amount = replace(panel_room_amount, "_", "")
			panel_garage_amount = replace(panel_garage_amount, "_", "")
			panel_subsidy_amount = replace(panel_subsidy_amount, "_", "")

			If panel_rent_amount <> "" Then rent_amount = rent_amount + panel_rent_amount
			If panel_lot_rent_amount <> "" Then lot_rent_amount = lot_rent_amount + panel_lot_rent_amount
			If panel_mortgage_amount <> "" Then mortgage_amount = mortgage_amount + panel_mortgage_amount
			If panel_insurance_amount <> "" Then insurance_amount = insurance_amount + panel_insurance_amount
			If panel_tax_amount <> "" Then tax_amount = tax_amount + panel_tax_amount
			If panel_room_amount <> "" Then room_amount = room_amount + panel_room_amount
			If panel_garage_amount <> "" Then garage_amount = garage_amount + panel_garage_amount
			If panel_subsidy_amount <> "" Then subsidy_amount = subsidy_amount + panel_subsidy_amount

			determined_shel = rent_amount + lot_rent_amount + mortgage_amount + insurance_amount + tax_amount + room_amount + garage_amount
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "CASH" Then
			cash_amount_yn = "Yes"

			EMReadscreen panel_cash_amount, 8, 8, 39
			panel_cash_amount = trim(panel_cash_amount)
			cash_amount = cash_amount + panel_cash_amount
			determined_assets = determined_assets + panel_cash_amount
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "ACCT" Then
			bank_account_yn = "Yes"

			ReDim Preserve ACCOUNTS_ARRAY(account_notes_const, acct_incrementor)

			EMReadscreen panel_acct_owner, 40, 4, 37
			EMReadscreen panel_acct_type, 2, 6, 44
			EMReadscreen panel_acct_bank, 20, 8, 44
			EMReadscreen panel_acct_amount, 8, 10, 46

			panel_acct_owner = trim(panel_acct_owner)
			name_array = ""
			name_array = split(panel_acct_owner, ", ")
			'TESTING CODE'
			' for name_test = 0 to UBound(name_array)
			' 	Msgbox "COUNTER - " & name_test & vbCr & "NAME - " & name_array(name_test)
			' Next
			panel_acct_owner = name_array(1) & " " & name_array(0)

			If panel_acct_type = "CK" Then ACCOUNTS_ARRAY(account_type_const, acct_incrementor) = "Checking"
			If panel_acct_type = "SV" Then ACCOUNTS_ARRAY(account_type_const, acct_incrementor) = "Savings"
			If panel_acct_type <> "SV" AND panel_acct_type <> "CK" Then ACCOUNTS_ARRAY(account_type_const, acct_incrementor) = "Other"
			ACCOUNTS_ARRAY(account_owner_const, acct_incrementor) = trim(panel_acct_owner)
			ACCOUNTS_ARRAY(bank_name_const, acct_incrementor) = replace(panel_acct_bank, "_", "")
			ACCOUNTS_ARRAY(account_amount_const, acct_incrementor) = trim(panel_acct_amount)

			determined_assets = determined_assets + ACCOUNTS_ARRAY(account_amount_const, acct_incrementor)
			ACCOUNTS_ARRAY(account_amount_const, acct_incrementor) = ACCOUNTS_ARRAY(account_amount_const, acct_incrementor) & ""
			acct_incrementor = acct_incrementor + 1
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "FACI" Then

		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "JOBS" Then
			jobs_income_yn = "Yes"
			ReDim Preserve JOBS_ARRAY(jobs_notes_const, jobs_incrementor)

			EMReadScreen panel_employee, 40, 4, 36
			EMReadScreen panel_employer, 30, 7, 42
			EMReadScreen panel_main_wage, 6, 6, 75
			panel_main_wage = replace(panel_main_wage, "_", "")
			panel_main_wage = trim(panel_main_wage)

			EMWriteScreen "X", 19, 38
			transmit
			EMReadScreen panel_frequency, 1, 5, 64
			EMReadScreen panel_hours, 6, 8, 64
			EMReadScreen panel_wage, 8, 9, 66
			EMReadScreen panel_monthly_income, 8, 18, 56

			panel_wage = replace(panel_wage, "_", "")
			panel_hours = replace(panel_hours, "_", "")
			panel_wage = trim(panel_wage)
			panel_hours = trim(panel_hours)
			panel_monthly_income = trim(panel_monthly_income)

			If panel_wage = "" Then
				If panel_main_wage <> "" Then
					panel_wage = panel_main_wage
				Else
					EMReadScreen panel_ave_wage, 8, 17, 56
					panel_ave_wage = trim(panel_ave_wage)
					If IsNumeric(panel_ave_wage) = True Then
						If panel_frequency = "1" Then panel_wage = panel_ave_wage/4.3
						If panel_frequency = "2" Then panel_wage = panel_ave_wage/2.15
						If panel_frequency = "3" Then panel_wage = panel_ave_wage/2
						If panel_frequency = "4" Then panel_wage = panel_ave_wage
					End If
				End If
			End If
			If panel_hours = "" Then
				EMReadScreen panel_ave_hours, 7, 16, 50
				panel_ave_hours = trim(panel_ave_hours)
				If IsNumeric(panel_ave_hours) = True Then
					MsgBox panel_ave_hours
					If panel_frequency = "1" Then panel_hours = panel_ave_hours/4.3
					If panel_frequency = "2" Then panel_hours = panel_ave_hours/2.15
					If panel_frequency = "3" Then panel_hours = panel_ave_hours/2
					If panel_frequency = "4" Then panel_hours = panel_ave_hours

				End If
			End If
			transmit

			panel_employee = trim(panel_employee)
			name_array = ""
			name_array = split(panel_employee, ", ")
			panel_employee = name_array(1) & " " & name_array(0)

			JOBS_ARRAY(jobs_employee_const, jobs_incrementor) = panel_employee
			JOBS_ARRAY(jobs_employer_const, jobs_incrementor) = replace(panel_employer, "_", "")
			JOBS_ARRAY(jobs_wage_const, jobs_incrementor) = panel_wage
			JOBS_ARRAY(jobs_hours_const, jobs_incrementor) = panel_hours
			If panel_frequency = "1" Then JOBS_ARRAY(jobs_frequency_const, jobs_incrementor) = "Monthly"
			If panel_frequency = "2" Then JOBS_ARRAY(jobs_frequency_const, jobs_incrementor) = "Semi-Monthly"
			If panel_frequency = "3" Then JOBS_ARRAY(jobs_frequency_const, jobs_incrementor) = "Biweekly"
			If panel_frequency = "4" Then JOBS_ARRAY(jobs_frequency_const, jobs_incrementor) = "Weekly"
			JOBS_ARRAY(jobs_monthly_pay_const, jobs_incrementor) = trim(panel_monthly_income)

			If IsNumeric(JOBS_ARRAY(jobs_monthly_pay_const, jobs_incrementor)) = True Then determined_income = determined_income + JOBS_ARRAY(jobs_monthly_pay_const, jobs_incrementor)
			jobs_incrementor = jobs_incrementor + 1
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "BUSI" Then
			busi_income_yn = "Yes"
			ReDim Preserve BUSI_ARRAY(busi_notes_const, busi_incrementor)

			EMReadScreen panel_owner, 25, 4, 37
			EMReadScreen panel_busi_type, 2, 5, 37
			EMReadScreen panel_monthly_wage, 8, 10, 69

			panel_owner = trim(panel_owner)
			name_array = ""
			name_array = split(panel_owner, ", ")
			panel_owner = name_array(1) & " " & name_array(0)

			BUSI_ARRAY(busi_owner_const, busi_incrementor) = panel_owner
			If panel_busi_type = "01" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Farming"
			If panel_busi_type = "02" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Real Estate"
			If panel_busi_type = "03" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Home Product Sales"
			If panel_busi_type = "04" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Sales"
			If panel_busi_type = "05" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Personal Services"
			If panel_busi_type = "06" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Paper Route"
			If panel_busi_type = "07" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "In Home Daycare"
			If panel_busi_type = "08" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Rental Income"
			If panel_busi_type = "09" Then BUSI_ARRAY(busi_info_const, busi_incrementor) = "Other"
			BUSI_ARRAY(busi_monthly_earnings_const, busi_incrementor) = trim(panel_monthly_wage)

			If IsNumeric(BUSI_ARRAY(busi_monthly_earnings_const, busi_incrementor)) = True Then
				determined_income = determined_income + BUSI_ARRAY(busi_monthly_earnings_const, busi_incrementor)
				BUSI_ARRAY(busi_annual_earnings_const, busi_incrementor) = FormatNumber(BUSI_ARRAY(busi_monthly_earnings_const, busi_incrementor) * 12, 2, -1, 0, -1)
			End If
			busi_incrementor = busi_incrementor + 1
		End If
		If PANELS_TO_READ_ARRAY(panel_type_const, each_panel) = "UNEA" Then
			unea_income_yn = "Yes"
			ReDim Preserve UNEA_ARRAY(unea_notes_const, unea_incrementor)

			EMReadScreen panel_earner, 25, 4, 36
			EMReadScreen panel_unea_type, 2, 5, 37

			EMWriteScreen "X", 10, 26
			transmit
			EMReadScreen panel_unea_frequency, 1, 5, 64
			EMReadScreen panel_unea_monthly_wage, 8, 18, 56
			If panel_unea_frequency = "4" Then
				EMReadScreen panel_unea_weekly_wage, 8, 17, 56
				UNEA_ARRAY(unea_weekly_earnings_const, unea_incrementor) = trim(panel_unea_weekly_wage)
			End If
			transmit

			panel_earner = trim(panel_earner)
			name_array = ""
			name_array = split(panel_earner, ", ")
			panel_earner = name_array(1) & " " & name_array(0)

			If panel_unea_type = "06" Then panel_unea_type = "Public Assistance not in MN"
			If panel_unea_type = "14" Then panel_unea_type = "Unemployment Insurance"
			If panel_unea_type = "19" or panel_unea_type = "21" OR panel_unea_type = "20" or panel_unea_type = "22" Then panel_unea_type = "Foster Care"
			If panel_unea_type = "16" Then panel_unea_type = "Railroad Retirement"
			If panel_unea_type = "17" Then panel_unea_type = "Retirement"
			If panel_unea_type = "35" or panel_unea_type = "37" or panel_unea_type = "40" Then panel_unea_type = "Spousal Support"
			If panel_unea_type = "18" Then panel_unea_type = "Military Entitlement"
			If panel_unea_type = "23" Then panel_unea_type = "Dividends"
			If panel_unea_type = "24" Then panel_unea_type = "Interest"
			If panel_unea_type = "25" Then panel_unea_type = "Prizes and Gifts"
			If panel_unea_type = "26" Then panel_unea_type = "Strike Benefit"
			If panel_unea_type = "27" Then panel_unea_type = "Contract for Deed"
			If panel_unea_type = "28" Then panel_unea_type = "Illegal Income"
			If panel_unea_type = "29" Then panel_unea_type = "Other Countable"
			If panel_unea_type = "30" Then panel_unea_type = "Infreq Irreg"
			If panel_unea_type = "31" Then panel_unea_type = "Other FS Only"
			If panel_unea_type = "45" Then panel_unea_type = "County 88 Gaming"
			If panel_unea_type = "47" Then panel_unea_type = "Tribal Income"
			If panel_unea_type = "48" Then panel_unea_type = "Trust Income"
			If panel_unea_type = "49" Then panel_unea_type = "Non-Recurring"
			If IsNumeric(panel_unea_type) = True Then panel_unea_type = ""

			UNEA_ARRAY(unea_owner_const, unea_incrementor) = panel_earner
			UNEA_ARRAY(unea_info_const, unea_incrementor) = panel_unea_type
			UNEA_ARRAY(unea_monthly_earnings_const, unea_incrementor) = trim(panel_unea_monthly_wage)

			If IsNumeric(UNEA_ARRAY(unea_monthly_earnings_const, unea_incrementor)) = True Then determined_income = determined_income + UNEA_ARRAY(unea_monthly_earnings_const, unea_incrementor)
			unea_incrementor = unea_incrementor + 1
		End If
		' MsgBox "PANEL - " & PANELS_TO_READ_ARRAY(panel_type_const, each_panel) & "-" & PANELS_TO_READ_ARRAY(panel_memb_const, each_panel) & "-" & PANELS_TO_READ_ARRAY(panel_inst_const, each_panel)
	Next


	determined_income = determined_income & ""
	determined_assets = determined_assets & ""
	determined_shel = determined_shel & ""

End If

show_pg_amounts = 1
show_pg_determination = 2
show_pg_review = 3

page_display = show_pg_amounts


Do
	Do
		err_msg = ""
		If page_display = show_pg_determination Then Call determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)
		If page_display = show_pg_review Then Call determine_actions(case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, is_elig_XFS, snap_denial_date, approval_date, date_of_application, do_we_have_applicant_id, action_due_to_out_of_state_benefits, mn_elig_begin_date, other_snap_state, case_has_previously_postponed_verifs_that_prevent_exp_snap, delay_action_due_to_faci, deny_snap_due_to_faci)

		If determined_utilities = "" Then determined_utilities = 0
		determined_income = FormatNumber(determined_income, 2, -1, 0, -1) & ""
		determined_assets = FormatNumber(determined_assets, 2, -1, 0, -1) & ""
		determined_shel = FormatNumber(determined_shel, 2, -1, 0, -1) & ""
		determined_utilities = FormatNumber(determined_utilities, 2, -1, 0, -1)
		calculated_resources = FormatNumber(calculated_resources, 2, -1, 0, -1)
		calculated_expenses = FormatNumber(calculated_expenses, 2, -1, 0, -1)

		BeginDialog Dialog1, 0, 0, 555, 385, "Full Expedited Determination"
		  ButtonGroup ButtonPressed
		  	If page_display = show_pg_amounts then
				Text 504, 12, 65, 10, "Amounts"

				GroupBox 5, 5, 390, 75, "Expedited Screening"
				If exp_screening_note_found = True Then
					Text 10, 20, 145, 10, "Information pulled from previous case note."
					Text 20, 35, 70, 10, "Income from CAF1: $ "
					Text 100, 35, 80, 10, caf_one_income
					Text 195, 35, 65, 10, "Assets from CAF1: $ "
					Text 270, 35, 75, 10, caf_one_assets
					Text 20, 50, 90, 10, "Housing from CAF1: $ "
					Text 100, 50, 65, 10, caf_one_rent
					Text 195, 50, 65, 10, "Utilities from CAF1: $ "
					Text 270, 50, 75, 10, caf_one_utilities
					Text 15, 65, 160, 10, xfs_screening
				End If
				If exp_screening_note_found = False Then
					Text 10, 20, 350, 10, "CASE:NOTE for Expedited Screening could not be found. No information to Display."
					Text 10, 30, 350, 10, "Review Application for screening answers"
				End If
				Text 10, 90, 370, 15, "Review and update the INCOME, ASSETS, and HOUSING EXPENSES as determined in the Interview."
				GroupBox 5, 105, 390, 125, "Information about Income, Resources, and Expenses"
				Text 15, 125, 60, 10, "Gross Income:    $"
				EditBox 75, 120, 155, 15, determined_income
				Text 15, 145, 35, 10, "Assets:   $"
				EditBox 50, 140, 180, 15, determined_assets
				Text 15, 165, 70, 10, "Shelter Expense:    $"
				EditBox 85, 160, 145, 15, determined_shel
				Text 15, 185, 60, 10, "Utilities Expense:"
				Text 77, 185, 145, 15, "$  " & determined_utilities
				PushButton 255, 120, 120, 13, "Calculate Income", income_calc_btn
				PushButton 255, 140, 120, 13, "Calculate Assets", asset_calc_btn
				PushButton 255, 160, 120, 13, "Calculate Housing Cost", housing_calc_btn
			    PushButton 255, 180, 120, 13, "Calculate Utilities", utility_calc_btn
				If snap_elig_results_read = True Then Text 55, 200, 180, 10, "Autofilled information based on current STAT and ELIG panels"
				Text 15, 215, 250, 10, "Blank amounts will be defaulted to ZERO."
				' GroupBox 5, 220, 390, 100, "Supports"
				' Text 15, 235, 260, 10, "If you need support in handling for expedited, please access these resources:"
			    ' PushButton 25, 250, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
				' PushButton 25, 265, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
				' PushButton 25, 280, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
				' PushButton 25, 295, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
			    ' PushButton 180, 250, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			    ' PushButton 180, 265, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
			    ' PushButton 180, 280, 150, 13, "CM 04.06 - 1st Mont Processing", cm_04_06_btn
			End If
			If page_display = show_pg_determination then
				Text 495, 27, 65, 10, "Determination"

				If is_elig_XFS = True Then Text 0, 25, 400, 10, "---------------------------------------------- This case IS EXPEDITED based on this critera: "
				If is_elig_XFS = False Then Text 0, 25, 400, 10, "---------------------------------------------- This case is NOT expedited based on this critera: "

				GroupBox 5, 5, 470, 135, "Expedited Determination"
				Text 15, 50, 120, 10, "Determination Amounts Entered:"
				Text 130, 50, 85, 10, "Total App Month Income:"
				Text 220, 50, 40, 10, "$ " & determined_income
				Text 130, 60, 85, 10, "Total App Month Assets:"
				Text 220, 60, 40, 10, "$ " & determined_assets
				Text 130, 70, 85, 10, "Total App Month Housing:"
				Text 220, 70, 40, 10, "$ " & determined_shel
				Text 130, 80, 85, 10, "Total App Month Utility:"
				Text 220, 80, 40, 10, "$ " & determined_utilities
				Text 295, 50, 135, 10, "Combined Resources (Income + Assets):"
				Text 430, 50, 40, 10, "$ " & calculated_resources
				Text 330, 70, 100, 10, "Combined Housing Expense:"
				Text 430, 70, 40, 10, "$ " & calculated_expenses

				GroupBox 5, 15, 470, 25, ""

				Text 295, 95, 125, 20, "Unit has less than $150 monthly Gross Income AND $100 or less in assets:"
				Text 430, 100, 35, 10, calculated_low_income_asset_test
				Text 295, 115, 125, 20, "Unit's combined resources are less than housing expense:"
				Text 430, 120, 35, 10, calculated_resources_less_than_expenses_test

				Text 18, 90, 65, 10, "Date of Application:"
				Text 85, 90, 50, 10, date_of_application
				Text 25, 100, 60, 10, "Date of Interview:"
				Text 85, 100, 50, 10, interview_date
				Text 25, 115, 60, 10, "Date of Approval:"
				EditBox 85, 110, 60, 15, approval_date
				Text 85, 125, 75, 10, "(or planned approval)"

				GroupBox 5, 135, 470, 155, "Possible Approval Delays"
			    Text 95, 150, 205, 10, "Is there a document for proof of identity of the applicant on file?"
			    DropListBox 300, 145, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", applicant_id_on_file_yn
			    Text 95, 165, 200, 10, "Can the Identity of the applicant be cleard through SOLQ/SMI?"
			    DropListBox 300, 160, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", applicant_id_through_SOLQ
			    PushButton 350, 160, 120, 13, "HOT TOPIC - Using SOLQ for ID", ht_id_in_solq_btn
			    Text 10, 185, 85, 10, "Explain Approval Delays:"
			    EditBox 95, 180, 375, 15, delay_explanation
			    Text 175, 205, 80, 10, "Specifc case situations:"
			    PushButton 255, 200, 215, 15, "SNAP is Active in Another State in " & MAXIS_footer_month & "/" & MAXIS_footer_year, snap_active_in_another_state_btn
			    PushButton 255, 215, 215, 15, "Expedited Approved Previously with Postponed Verifications", case_previously_had_postponed_verifs_btn
			    PushButton 255, 230, 215, 15, "Household is Currently in a Facility", household_in_a_facility_btn
				Text 15, 255, 330, 10, "If it is already determined that SNAP should be denied, enter a denial date and explanation of denial."
				Text 355, 255, 65, 10, "SNAP Denial Date:"
				EditBox 420, 250, 50, 15, snap_denial_date
				Text 30, 275, 65, 10, "Denial Explanation:"
				EditBox 95, 270, 375, 15, snap_denial_explain
			End If
			If page_display = show_pg_review then
				Text 507, 42, 65, 10, "Review"

				GroupBox 5, 5, 470, 115, "Actions to Take"
				Text 20, 30, 45, 10, "Next Steps:"

				Text 15, 20, 280, 10, case_assesment_text

				Text 25, 40, 435, 20, next_steps_one
				Text 25, 60, 435, 20, next_steps_two
				Text 25, 80, 435, 20, next_steps_three
				Text 25, 100, 435, 20, next_steps_four

				EditBox 800, 800, 50, 15, fake_box_that_does_nothing
				Text 310, 15, 100, 10, "For help with the next steps:"
				PushButton 310, 25, 155, 13, "Request Support from Knowledge Now", knowledge_now_support_btn

				GroupBox 5, 120, 470, 85, "Postponed Verifications"
				If is_elig_XFS = True AND IsDate(snap_denial_date) = False Then
					Text 15, 135, 160, 10, "Are there Postponed Verifications for this case?"
					DropListBox 180, 130, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", postponed_verifs_yn
					Text 20, 155, 80, 10, "Postponed Verifications:"
					EditBox 105, 150, 360, 15, list_postponed_verifs
				    PushButton 320, 130, 145, 13, "TE 02.10.01 EXP w/ Pending Verifs", te_02_10_01_btn
				    Text 20, 175, 120, 10, "Can I postpone Verifications for ..."
				    Text 145, 175, 70, 10, "Immigration - YES."
				    Text 225, 175, 55, 10, "Sponsor - YES."
				    Text 300, 175, 125, 10, "anything OTHER than ID - YES. "
					Text 30, 190, 300, 10, "Appplicant's identity is the ONLY required verification to approve Expedited SNAP."
				    PushButton 320, 187, 145, 13, "CM 04.12 Verification Requirement for EXP", cm_04_12_btn
				End If
				If is_elig_XFS = False Then
					Text 15, 135, 450, 10, "We cannot postpone any verifications for a case that does not meet Expedited criteria."
				End If
				If IsDate(snap_denial_date) = True Then
					Text 15, 135, 450, 10, "Additional verifications are not needed if a Denial has already been determined."
				End If

			    GroupBox 5, 205, 470, 70, "EBT Information"
				If IsDate(snap_denial_date) = True Then
					Text 15, 220, 415, 10, "Advise resident to keep track of an EBT card they have received, even though the application is being denied."
					Text 20, 235, 415, 10, "If the case ever reapplies, or is determined eligible, the EBT card remains connected to the case and getting benefits will be easier."
				Else
					Text 15, 220, 335, 10, "Do not delay in approving SNAP benefits due to if the household does or does not have an EBT card."
				    Text 20, 235, 415, 10, "If there has never been a card issued for a case, approving the benefit with an REI will prevent a card from being sent via mail."
				    Text 20, 245, 305, 10, "If a case needs the first card mailed, do NOT REI benefits as they will not receive their card."
				End If
				Text 15, 260, 255, 10, "EBT Card issues can be complicated. Refer to the EBT Card Information here:"
			    PushButton 270, 257, 195, 13, "Temporary Program Changes - EBT Cards ", temp_prog_changes_ebt_card_btn

			End If
			GroupBox 5, 295, 470, 60, "If you need support in handling for expedited, please access these resources:"
			PushButton 15, 305, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
			PushButton 15, 320, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
			PushButton 15, 335, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			PushButton 165, 305, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
			PushButton 165, 320, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
			PushButton 315, 305, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
			PushButton 315, 320, 150, 13, "CM 04.06 - 1st Mont Processing", cm_04_06_btn

		    If page_display <> show_pg_amounts then PushButton 485, 10, 65, 13, "Amounts", amounts_btn
		    If page_display <> show_pg_determination then PushButton 485, 25, 65, 13, "Determination", determination_btn
		    If page_display <> show_pg_review then PushButton 485, 40, 65, 13, "Review", review_btn
		    If page_display <> show_pg_review then PushButton 445, 365, 50, 15, "Next", next_btn
			If page_display = show_pg_review then PushButton 445, 365, 50, 15, "Finish", finish_btn
		    CancelButton 500, 365, 50, 15
		    ' OkButton 500, 350, 50, 15
		EndDialog

		Dialog Dialog1
		cancel_confirmation
		' MsgBox "1 - ButtonPressed is " & ButtonPressed

		If ButtonPressed = -1 Then
			If page_display <> show_pg_review then ButtonPressed = next_btn
			If page_display = show_pg_review then ButtonPressed = finish_btn
		End If

		If ButtonPressed = income_calc_btn Then Call app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, JOBS_ARRAY, BUSI_ARRAY, UNEA_ARRAY)
		If ButtonPressed = asset_calc_btn Then Call app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, ACCOUNTS_ARRAY)
		If ButtonPressed = housing_calc_btn Then Call app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)
		If ButtonPressed = utility_calc_btn Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
		If ButtonPressed = snap_active_in_another_state_btn Then
			If IsDate(date_of_application) = False Then MsgBox "Attention:" & vbCr & vbCr & "The funcationality to determine actions if a household is reporting benefits in another state cannot be run if a valid application date has not been entered."
			If IsDate(date_of_application) = True Then Call snap_in_another_state_detail(date_of_application, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)
		End If
		If ButtonPressed = case_previously_had_postponed_verifs_btn Then Call previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
		If ButtonPressed = household_in_a_facility_btn Then Call household_in_a_facility_detail(delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, delay_explanation, snap_denial_explain, snap_denial_date, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)

		If ButtonPressed = knowledge_now_support_btn Then Call send_support_email_to_KN
		If ButtonPressed = te_02_10_01_btn Then Call view_poli_temp("02", "10", "01", "")

		' MsgBox "2 - ButtonPressed is " & ButtonPressed

		' If page_display = show_pg_amounts Then
		'
		' End If
		If page_display = show_pg_determination Then
			delay_due_to_interview = False
			do_we_have_applicant_id = "UNKNOWN"
			If applicant_id_on_file_yn = "Yes" OR applicant_id_through_SOLQ = "Yes" Then do_we_have_applicant_id = True
			If applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "No" Then do_we_have_applicant_id = False

			' If IsDate(date_of_application) = False Then err_msg = err_msg & vbCr & "* The date of application needs to be entered as a valid date."
			' If IsDate(interview_date) = False Then err_msg = err_msg & vbCr & "* The interview date needs to be entered as a valid date. An Expedited Determination cannot be completed without the interview."
			If IsDate(snap_denial_date) = True Then
				If DateDiff("d", date, snap_denial_date) > 0 Then err_msg = err_msg & vbCr & "* Future Date denials or 'Possible' denials are not what the 'SNAP Denial Date' field is for." & vbCr &_
																						  "* Only indicate a denial if you already have enough information to determine that the SNAP application should be denied." & vbCr &_
																						  "* If this is the determination, review the date in the SNAP Denial Field as it appears to be a future date."
				snap_denial_explain = trim(snap_denial_explain)
				If len(snap_denial_explain) < 10 then err_msg = err_msg & vbCr & "* Since this SNAP case is to be denied, explain the reason for denial in detail."
			Else
				If is_elig_XFS = True Then
					If IsDate(approval_date) = True Then
						If DateDiff("d", date, approval_date) > 0 Then err_msg = err_msg & vbCr & "* Approvals should happen the same day an Expedited Determination is completed if the case is Expedited. Since the Income, Assets, and Expenses indicate this case is expedited AND we appear to be ready to approve, this should be completed today."
						' If DateDiff("d", interview_date, date) < 0 Then
					End If
					If applicant_id_on_file_yn = "?" AND applicant_id_through_SOLQ = "?" Then
						err_msg = err_msg & vbCr & "* Indicate if we have identity of the applicant on file or available through SOLQ"
					ElseIf applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "?" Then
						err_msg = err_msg & vbCr & "* Since there is no identity found in the file for the applicant, check SOLQ/SMI to verify identity."
					ElseIf applicant_id_on_file_yn = "?" AND applicant_id_through_SOLQ = "No" Then
						err_msg = err_msg & vbCr & "* Since the applicant's identity cannot be cleared through SOLQ/SMI, check the case file and person file for documents that can be used to verify identity. Remember that SNAP does NOT require a Photo ID or Official Government ID."
					End If

					'Defaulting Delay Explanation
					If IsDate(approval_date) = True AND IsDate(interview_date) = True AND IsDate(date_of_application) = True Then
						If DateDiff("d", date_of_application, approval_date) > 7 Then
							If DateDiff("d", interview_date, approval_date) = 0 Then delay_due_to_interview = True
						End If
					End If
					If delay_due_to_interview = True AND InStr(delay_explanation, "Approval of Expedited delayed until completion of Interview") = 0 Then
						delay_explanation = delay_explanation & "; Approval of Expedited delayed until completion of Interview."
					End If
					If delay_due_to_interview = False then
						delay_explanation = replace(delay_explanation, "Approval of Expedited delayed until completion of Interview.", "")
						delay_explanation = replace(delay_explanation, "Approval of Expedited delayed until completion of Interview", "")
					End If
					If do_we_have_applicant_id = False AND InStr(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant") = 0 Then
						delay_explanation = delay_explanation & "; Approval cannot be completed as we have NO Proof of Identity for the Applicant."
					End If
					If do_we_have_applicant_id <> False Then
						delay_explanation = replace(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant.", "")
						delay_explanation = replace(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant", "")
					End If

					Call format_explanation_text(delay_explanation)
					Call format_explanation_text(snap_denial_explain)

					expedited_approval_delayed = False
					If IsDate(approval_date) = False Then expedited_approval_delayed = True
					If IsDate(approval_date) = True  AND IsDate(date_of_application) = True Then
						If DateDiff("d", date_of_application, approval_date) > 7 Then expedited_approval_delayed = True
					End If
					If expedited_approval_delayed = True AND len(delay_explanation) < 20 Then err_msg = err_msg & vbCR & "* The approval of the Expedited SNAP is or has been delayed. Provide a detailed explaination of the reason for delay or complete the approval."

				End If
				If is_elig_XFS = False Then

				End If
			End If

		End If
		If page_display = show_pg_review Then
			If postponed_verifs_yn = "Yes" AND trim(list_postponed_verifs) = "" Then err_msg = err_msg & vbCr & "* Since you have Postponed Verifications indicated, list what they are for the NOTE."
		End If

		' MsgBox "3 - ButtonPressed is " & ButtonPressed


		If ButtonPressed = next_btn AND err_msg = "" Then page_display = page_display + 1
		If ButtonPressed = amounts_btn Then page_display = show_pg_amounts
		If ButtonPressed = determination_btn AND err_msg = "" Then page_display = show_pg_determination
		If ButtonPressed = review_btn AND err_msg = "" AND page_display <> show_pg_amounts Then page_display = show_pg_review
		If ButtonPressed = review_btn AND err_msg = "" AND page_display = show_pg_amounts Then page_display = show_pg_determination

		If err_msg <> "" And ButtonPressed < 100 AND page_display <> show_pg_amounts Then MsgBox "***** Action Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg

		If ButtonPressed <> finish_btn Then err_msg = "LOOP"
		' MsgBox "4 - ButtonPressed is " & ButtonPressed

		If ButtonPressed >= 1000 Then
			If ButtonPressed = hsr_manual_expedited_snap_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Expedited_SNAP.aspx"
			If ButtonPressed = hsr_snap_applications_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/SNAP_Applications.aspx"
			If ButtonPressed = ryb_exp_identity_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%201%20-%20Identity.mp4"
			If ButtonPressed = ryb_exp_timeliness_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%202%20-%20Timeliness.mp4"
			If ButtonPressed = sir_exp_flowchart_btn Then resource_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/Documents/SNAP%20Expedited%20Service%20Flowchart.pdf"
			If ButtonPressed = cm_04_04_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000404"
			If ButtonPressed = cm_04_06_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000406"
			If ButtonPressed = ht_id_in_solq_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/How-to-use-SMI-SOLQ-to-verify-ID-for-SNAP.aspx"
			If ButtonPressed = cm_04_12_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000412"
			If ButtonPressed = temp_prog_changes_ebt_card_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Temporary-Program-Changes--EBT-cards,-checks,-bus-cards.aspx"

			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & resource_URL
		End If



	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If developer_mode = False Then

	txt_file_name = "expedited_determination_detail_" & MAXIS_case_number & "_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".txt"
	exp_info_file_path = t_drive &"\Eligibility Support\Assignments\Expedited Information\"  & txt_file_name
	' MsgBox exp_info_file_path

	With objFSO
		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		Set objTextStream = .OpenTextFile(exp_info_file_path, ForWriting, true)

		objTextStream.WriteLine ""

		objTextStream.WriteLine "CASE NUMBER ^*^*^" & MAXIS_case_number
		objTextStream.WriteLine "WORKER NAME ^*^*^" & worker_name
		objTextStream.WriteLine "CASE X NUMBER  ^*^*^" & case_pw
		If IsDate(date_of_application) = True Then date_of_application = DateAdd("d", 0, date_of_application)
		objTextStream.WriteLine "DATE OF APPLICATION ^*^*^" & date_of_application
		If IsDate(appt_notc_sent_on) = True Then appt_notc_sent_on = DateAdd("d", 0, appt_notc_sent_on)
		objTextStream.WriteLine "APPT NOTC SENT DATE ^*^*^" & appt_notc_sent_on
		If IsDate(appt_date_in_note) = True Then appt_date_in_note = DateAdd("d", 0, appt_date_in_note)
		objTextStream.WriteLine "APPT DATE ^*^*^" & appt_date_in_note
		If IsDate(interview_date) = True Then interview_date = DateAdd("d", 0, interview_date)
		objTextStream.WriteLine "DATE OF INTERVIEW ^*^*^" & interview_date
		objTextStream.WriteLine "EXPEDITED SCREENING STATUS ^*^*^" & xfs_screening
		objTextStream.WriteLine "EXPEDITED DETERMINATION STATUS ^*^*^" & is_elig_XFS
		objTextStream.WriteLine "DET INCOME ^*^*^" & determined_income
		objTextStream.WriteLine "DET ASSETS ^*^*^" & determined_assets
		objTextStream.WriteLine "DET SHEL ^*^*^" & determined_shel
		objTextStream.WriteLine "DET HEST ^*^*^" & determined_utilities
		If IsDate(approval_date) = True Then approval_date = DateAdd("d", 0, approval_date)
		objTextStream.WriteLine "DATE OF APPROVAL ^*^*^" & approval_date
		If IsDate(snap_denial_date) = True Then snap_denial_date = DateAdd("d", 0, snap_denial_date)
		objTextStream.WriteLine "SNAP DENIAL DATE ^*^*^" & snap_denial_date
		objTextStream.WriteLine "SNAP DENIAL REASON ^*^*^" & snap_denial_explain
		objTextStream.WriteLine "ID ON FILE ^*^*^" & do_we_have_applicant_id
		objTextStream.WriteLine "OUTSTATE ACTION ^*^*^" & action_due_to_out_of_state_benefits
		objTextStream.WriteLine "OUTSTATE STATE ^*^*^" & other_snap_state
		If IsDate(other_state_reported_benefit_end_date) = True Then other_state_reported_benefit_end_date = DateAdd("d", 0, other_state_reported_benefit_end_date)
		objTextStream.WriteLine "OUTSTATE REPORTED END DATE ^*^*^" & other_state_reported_benefit_end_date
		objTextStream.WriteLine "OUTSTATE OPENENDED ^*^*^" & other_state_benefits_openended
		If IsDate(other_state_verified_benefit_end_date) = True Then other_state_verified_benefit_end_date = DateAdd("d", 0, other_state_verified_benefit_end_date)
		objTextStream.WriteLine "OUTSTATE VERIFIED END DATE ^*^*^" & other_state_verified_benefit_end_date
		If IsDate(mn_elig_begin_date) = True Then mn_elig_begin_date = DateAdd("d", 0, mn_elig_begin_date)
		objTextStream.WriteLine "MN ELIG BEGIN DATE ^*^*^" & mn_elig_begin_date
		objTextStream.WriteLine "PREV POST DELAY APP ^*^*^" & case_has_previously_postponed_verifs_that_prevent_exp_snap				'(Boolean)
		If IsDate(previous_date_of_application) = True Then previous_date_of_application = DateAdd("d", 0, previous_date_of_application)
		objTextStream.WriteLine "PREV POST PREV DATE OF APP ^*^*^" & previous_date_of_application
		objTextStream.WriteLine "PREV POST LIST ^*^*^" & prev_verif_list
		objTextStream.WriteLine "PREV POST CURR VERIF POST ^*^*^" & curr_verifs_postponed_yn
		objTextStream.WriteLine "PREV POST ONGOING SNAP APP ^*^*^" & ongoing_snap_approved_yn
		objTextStream.WriteLine "PREV POST VERIFS RECVD ^*^*^" & prev_post_verifs_recvd_yn
		objTextStream.WriteLine "EXPLAIN APPROVAL DELAYS  ^*^*^" & delay_explanation								'(all of them)
		objTextStream.WriteLine "POSTPONED VERIFICATIONS ^*^*^" & postponed_verifs_yn
		objTextStream.WriteLine "WHAT ARE THE POSTPONED VERIFICATIONS ^*^*^" & list_postponed_verifs
		objTextStream.WriteLine "FACI DELAY ACTION ^*^*^" & delay_action_due_to_faci
		objTextStream.WriteLine "FACI DENY ^*^*^" & deny_snap_due_to_faci
		objTextStream.WriteLine "FACI NAME ^*^*^" & facility_name
		objTextStream.WriteLine "FACI INELIG SNAP ^*^*^" & snap_inelig_faci_yn
		If IsDate(faci_entry_date) = True Then faci_entry_date = DateAdd("d", 0, faci_entry_date)
		objTextStream.WriteLine "FACI ENTRY DATE ^*^*^" & faci_entry_date
		If IsDate(faci_release_date) = True Then faci_release_date = DateAdd("d", 0, faci_release_date)
		objTextStream.WriteLine "FACI RELEASE DATE ^*^*^" & faci_release_date
		objTextStream.WriteLine "FACI RELEASE IN 30 DAYS ^*^*^" & release_within_30_days_yn
		objTextStream.WriteLine "DATE OF SCRIPT RUN ^*^*^" & now
        objTextStream.WriteLine "SCRIPT RUN ^*^*^EXPEDITED DETERMINATION"

		'Close the object so it can be opened again shortly
		objTextStream.Close

	End With

End if

'Commented out while we determine if this action is warranted
'Need to add functionality to NOT send it if a QI member is running the script.
' If delay_explanation <> "" AND is_elig_XFS = True AND IsDate(approval_date) = False Then
' 	email_subject = "Auto Generated Email to review EXPEDITED DELAY"
' 	email_body = "This email is automatically generated without request from the worker for a case that is determined Expedited and has a delay reason."
'
' 	email_body = email_body & vbCr & vbCr & "The case # " & MAXIS_case_number & " was determined to meet Expedited Criteria."
' 	email_body = email_body & vbCr & "Income: $ " & determined_income
' 	email_body = email_body & vbCr & "Assets: $ " & determined_assets
' 	email_body = email_body & vbCr & "Housing: $ " & determined_shel
' 	email_body = email_body & vbCr & "Utility: $ " & determined_utilities
' 	If do_we_have_applicant_id = False Then email_body = email_body & vbCr & vbCr & "According to the worker there is NO proof of identity on file."
' 	If do_we_have_applicant_id = True Then email_body = email_body & vbCr & vbCr & "According to the worker, we do have proof of identity on file."
' 	email_body = email_body & vbCr & vbCr & "Reasons cited for the delay:"
' 	email_body = email_body & vbCr & delay_explanation
'
' 	email_body = "~~This email is generated from within the 'Expedited Determination' Script.~~" & vbCr & vbCr & email_body
' 	call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", True)
' End If


note_calculation_detail = False
If income_review_completed = True OR assets_review_completed = True OR shel_review_completed = True Then note_calculation_detail = True

note_case_situation_details = False
If action_due_to_out_of_state_benefits <> "" OR prev_post_verif_assessment_done = True OR faci_review_completed = True Then note_case_situation_details = True

'creating a custom header: this is read by BULK - EXP SNAP REVIEW script so don't mess this please :)
If IsDate(snap_denial_date) = TRUE Then
	case_note_header_text = "Expedited Determination: SNAP to be denied"
Else
	IF is_elig_XFS = True then
		case_note_header_text = "Expedited Determination: SNAP appears expedited"
	ELSEIF is_elig_XFS = False then
		case_note_header_text = "Expedited Determination: SNAP does not appear expedited"
	END IF
End If

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------
navigate_to_MAXIS_screen "CASE", "NOTE"
Call start_a_blank_case_note
Call write_variable_in_case_note (case_note_header_text)
If interview_date <> "" Then Call write_variable_in_case_note (" - Interview completed on: " & interview_date & " and full Expedited Determination Done")
IF exp_screening_note_found = TRUE Then
	Call write_variable_in_case_note ("Expedited Screening found: " & xfs_screening)
	Call write_variable_in_case_note ("  Based on: Income:  $ " & right("        " & caf_one_income, 8) & ", Assets:    $ " & right("        " & caf_one_assets, 8)    & ", Totaling: $ " & right("        " & caf_one_resources, 8))
	Call write_variable_in_case_note ("            Shelter: $ " & right("        " & caf_one_rent, 8)   & ", Utilities: $ " & right("        " & caf_one_utilities, 8) & ", Totaling: $ " & right("        " & caf_one_expenses, 8))

	Call write_variable_in_case_note ("---")
End If
If IsDate(snap_denial_date) = TRUE Then
	Call write_variable_in_CASE_NOTE("SNAP to be denied on " & snap_denial_date & ". Since case is not SNAP eligible, case cannot receive Expedited issuance.")
	If is_elig_XFS = TRUE Then
		Call write_variable_with_indent_in_CASE_NOTE("Case is determined to meet criteria based upon income alone.")
		Call write_variable_with_indent_in_CASE_NOTE("Expedited approval requires case to be otherwise eligble for SNAP and this does not meet this criteria.")
	ElseIf is_elig_XFS = False Then
		Call write_variable_with_indent_in_CASE_NOTE("Expedited SNAP cannot be approved as case does not meet all criteria")
	End If
	Call write_bullet_and_variable_in_CASE_NOTE("Explanation of Denial", snap_denial_explain)
Else
	IF is_elig_XFS = TRUE Then
		Call write_variable_in_case_note ("Case is determined to meet criteria for Expedited SNAP.")
		If IsDate(approval_date) = False AND delay_explanation <> "" Then
			Call write_variable_in_case_note (" - Approval of Expedited SNAP cannot be completed due to:")
			' delay_explanation = THIS NEEDS TO BE AN ARRAY
			If InStr(delay_explanation, ";") = 0 Then
				delay_explain_array = Array(delay_explanation)
			Else
				delay_explain_array = Split(delay_explanation, ";")
			End If
			counter = 1
			For each item in delay_explain_array
				item = trim(item)
				Call write_variable_with_indent_in_CASE_NOTE(counter & ". " & item)
				counter = counter + 1
			Next
		End If
	End If
	IF is_elig_XFS = FALSE Then Call write_variable_in_case_note ("Case does not meet Expedited SNAP criteria.")
	Call write_variable_in_case_note ("  Based on: Income:  $ " & right("        " & determined_income, 8) & ", Assets:    $ " & right("        " & determined_assets, 8)   & ", Totaling: $ " & right("        " & calculated_resources, 8))
	Call write_variable_in_case_note ("            Shelter: $ " & right("        " & determined_shel, 8)   & ", Utilities: $ " & right("        " & determined_utilities, 8) & ", Totaling: $ " & right("        " & calculated_expenses, 8))
	Call write_variable_in_CASE_NOTE("  --- Expedited Criteria Tests ---")
	If calculated_low_income_asset_test = False Then Call write_variable_in_case_note("  FAILED - Resources Less than or Equal to $100 and Income Less than $150")
	If calculated_low_income_asset_test = True Then Call write_variable_in_case_note("  PASSED - Resources Less than or Equal to $100 and Income Less than $150")
	If calculated_resources_less_than_expenses_test = False Then Call write_variable_in_case_note("  FAILED - Resources Plus Income Less than Shelter Costs")
	If calculated_resources_less_than_expenses_test = True Then Call write_variable_in_case_note("  PASSED - Resources Plus Income Less than Shelter Costs")
	Call write_variable_in_case_note ("---")
	IF is_elig_XFS = TRUE Then
		Call write_variable_in_case_note ("Important Details")
		Call write_bullet_and_variable_in_case_note ("Date of Application", date_of_application)
		Call write_bullet_and_variable_in_case_note ("Date of Interview", interview_date)
		Call write_bullet_and_variable_in_case_note ("Date of Approval", approval_date)
		' Call write_bullet_and_variable_in_case_note ("Reason for Delay", delay_explanation)
		Call write_bullet_and_variable_in_CASE_NOTE("Postponed Verifs", list_postponed_verifs)
		Call write_variable_in_case_note ("---")
	End If
	If note_calculation_detail = True Then
		Call write_variable_in_case_note ("* Additional Notes about these amounts:")
		If income_review_completed = True Then
			' Call write_variable_in_case_note ("*   INCOME Details:")
			If jobs_income_yn = "Yes" Then
				' Call write_variable_in_case_note ("    - JOBS")
				for the_job = 0 to UBound(JOBS_ARRAY, 2)
					If IsNumeric(JOBS_ARRAY(jobs_wage_const, the_job)) = True AND IsNumeric(JOBS_ARRAY(jobs_hours_const, the_job)) = True Then
						Call write_variable_in_case_note ("  - JOBS: " & JOBS_ARRAY(jobs_employee_const, the_job) & " at " & JOBS_ARRAY(jobs_employer_const, the_job) & ": $" & JOBS_ARRAY(jobs_wage_const, the_job) & "/hr at " & JOBS_ARRAY(jobs_hours_const, the_job) & " hrs/wk.")
						Call write_variable_in_case_note ("            - Monthly Gross: $" & JOBS_ARRAY(jobs_monthly_pay_const, the_job))
					End If
				Next
			End If
			If busi_income_yn = "Yes" Then
				' Call write_variable_in_case_note ("    - SELF EMPLOYMENT")
				for the_busi = 0 to UBound(BUSI_ARRAY, 2)
					Call write_variable_in_case_note ("  - BUSI: " & BUSI_ARRAY(busi_owner_const, the_busi) & " for " & BUSI_ARRAY(busi_info_const, the_busi) & ".")
					Call write_variable_in_case_note ("            - Monthly Gross: $" & BUSI_ARRAY(busi_monthly_earnings_const, the_busi))
				Next
			End If
			If unea_income_yn = "Yes" Then
				' Call write_variable_in_case_note ("    - UNEARNED INCOME")
				for the_unea = 0 to UBound(UNEA_ARRAY, 2)
					Call write_variable_in_case_note ("  - UNEA: " & UNEA_ARRAY(unea_owner_const, the_unea) & " from " & UNEA_ARRAY(unea_info_const, the_unea) & ".")
					Call write_variable_in_case_note ("            - Monthly Gross: $" & UNEA_ARRAY(unea_monthly_earnings_const, the_unea))
				Next
			End If
			' app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, JOBS_ARRAY, BUSI_ARRAY, UNEA_ARRAY)
		End If
		If assets_review_completed = True Then
			' Call write_variable_in_case_note ("*   ASSET Details:")
			If cash_amount_yn = "Yes" Then Call write_variable_in_case_note ("  - CASH: Amount: $" & cash_amount)
			If bank_account_yn = "Yes" Then
				' Call write_variable_in_case_note ("    - BANK ACCOUNTS")
				For the_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
					If ACCOUNTS_ARRAY(account_type_const, the_acct) <> "Select One..." Then
						acct_info = "  - ACCT: " & ACCOUNTS_ARRAY(account_type_const, the_acct)
						If ACCOUNTS_ARRAY(bank_name_const, the_acct) <> "" Then acct_info = acct_info & " at " & ACCOUNTS_ARRAY(bank_name_const, the_acct)
						If ACCOUNTS_ARRAY(account_owner_const, the_acct) <> "" Then acct_info = acct_info & " owned by: " & ACCOUNTS_ARRAY(account_owner_const, the_acct)
						acct_info = acct_info & ". Balance: $" & ACCOUNTS_ARRAY(account_amount_const, the_acct)
						Call write_variable_in_case_note (acct_info)
					End If
				Next
			End If
			' app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, ACCOUNTS_ARRAY)
		End If
		If shel_review_completed = True Then
			' Call write_variable_in_case_note ("*   SHELTER Details:")
			first_housing_detail = True
			If rent_amount <> "" OR lot_rent_amount <> "" OR mortgage_amount <> "" OR insurance_amount <> "" OR tax_amount <> "" OR room_amount <> "" OR garage_amount <> "" Then

				Call write_variable_in_case_note ("  - SHEL: Rent:     $ " & right("    " & rent_amount, 4)    &  "   -   Lot Rent:  $" & right("    " & lot_rent_amount, 4))
				Call write_variable_in_case_note ("          Mortgage: $ " & right("    " & mortgage_amount, 4) & "   -   Insurance: $" & right("    " & insurance_amount, 4))
				Call write_variable_in_case_note ("          Tax:      $ " & right("    " & tax_amount, 4)      & "   -   Room:      $" & right("    " & room_amount, 4))
				Call write_variable_in_case_note ("          Garage:   $ " & right("    " & garage_amount, 4))
				Call write_variable_in_case_note ("          SUBSIDY:  $ " & right("    " & subsidy_amount, 4))
			End If
		End If
	End If
	' Call write_variable_in_case_note ("*   UTILITY Details:")
	If all_utilities <> "" Then Call write_variable_in_case_note ("  - HEST: " & all_utilities)
	' app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)

End If

If note_case_situation_details = True Then
	Call write_variable_in_case_note ("---")
	Call write_variable_in_case_note ("Additional details about this case:")

	If action_due_to_out_of_state_benefits <> "" Then Call write_variable_in_case_note ("* SNAP in Another State")
	If action_due_to_out_of_state_benefits = "DENY" Then
		Call write_variable_in_case_note ("*   SNAP to be DENIED as active in another state for the application processing 30 days.")
		If other_snap_state <> "" Then Call write_variable_in_case_note ("      - Other State: " & other_snap_state)
		Call write_variable_in_case_note ("      - Date of Application: " & date_of_application)
		Call write_variable_in_case_note ("      - Day 30: " & day_30_from_application)
		If IsDate(other_state_verified_benefit_end_date) = True  Then
			Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_verified_benefit_end_date & " - this date has been confirmed")
		ElseIF IsDate(other_state_reported_benefit_end_date) = True Then
			Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_reported_benefit_end_date & " - reported")
		End If
		' Call write_variable_in_case_note ("      - Date of Application: " & date_of_application)
	End If
	If action_due_to_out_of_state_benefits = "APPROVE" Then
		Call write_variable_in_case_note ("*   SNAP can be approved in MN for a later date.")
		If other_snap_state <> "" Then Call write_variable_in_case_note ("      - Other State: " & other_snap_state)
		Call write_variable_in_case_note ("      - Date of Application: " & date_of_application)
		Call write_variable_in_case_note ("      - Begin Date of Eligibility in MN: " & mn_elig_begin_date)
		Call write_variable_in_case_note ("      - Day 30: " & day_30_from_application)
		If IsDate(other_state_verified_benefit_end_date) = True  Then
			Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_verified_benefit_end_date & " - this date has been confirmed")
		ElseIF IsDate(other_state_reported_benefit_end_date) = True Then
			Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_reported_benefit_end_date & " - reported")
		End If
	End If
	If action_due_to_out_of_state_benefits = "FOLLOW UP" Then
		Call write_variable_in_case_note ("*   Needs response/additional information and is causing a delay in processing")
		If other_snap_state <> "" Then Call write_variable_in_case_note ("      - Other State: " & other_snap_state)
		Call write_variable_in_case_note ("      - The end date of benefits is open-ended or unknown and needs response from the other state before we can take action on the case in MN.")
	End If
		' snap_in_another_state_detail(date_of_application, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)

	If prev_post_verif_assessment_done = True Then
		Call write_variable_in_case_note ("* SNAP previously Approved with Postponed Verifciations")
		If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then
			eff_close_date = replace(previous_expedited_package, "/", "/1/")
			eff_close_date = DateAdd("m", 1, eff_close_date)
			eff_close_date = DateAdd("d", -1, eff_close_date)
			Call write_variable_in_case_note ("*   Expedited SNAP package cannot be approved due to unreceived postponed Verificactions")
			Call write_variable_in_case_note ("      - Previousl application on " & previous_date_of_application & " was approved as EXPEDITED with POSTPONED VERIFICATIONS.")
			Call write_variable_in_case_note ("      - This package closed on " & eff_close_date & ".")
			Call write_variable_in_case_note ("      - The postponed verifications have still not been received.")
			Call write_variable_in_case_note ("      - Previously postponed verifs: " & prev_verif_list)
			Call write_variable_in_case_note ("      - In order to approve the new Expedited Package for the current application, we would need to postpone verifications AGAIN.")
		End If
		If case_has_previously_postponed_verifs_that_prevent_exp_snap = False Then
			Call write_variable_in_case_note ("*   Though the case had previously postponed verifications, current Expedited can be approved")
			If prev_verifs_mandatory_yn = "No" Then Call write_variable_in_case_note ("      - The previous postponed verifications were not mandatory and case meet requirements for regular SNAP.")
			If curr_verifs_postponed_yn = "No" Then Call write_variable_in_case_note ("      - The current application does not require postponed verifications to be approved and case meet requirements for regular SNAP.")
			If ongoing_snap_approved_yn = "Yes" Then Call write_variable_in_case_note ("      - The case was approved for regular SNAP.")
			If prev_post_verifs_recvd_yn = "Yes" Then Call write_variable_in_case_note ("      - The previously postponed verifications have been received.")

		End If

		' previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
	End If
	If faci_review_completed = True Then
		If delay_action_due_to_faci = True Then
			Call write_variable_in_case_note ("* Resident is in a facility ")
			Call write_variable_in_case_note ("*  Expedited SNAP cannot be processed at this time.")
			If facility_name <> "" Then Call write_variable_in_case_note ("      - Facility Name: " & facility_name & " - an Ineligible SNAP Facility")
			If facility_name = "" Then Call write_variable_in_case_note ("      - Resident is in an Ineligible SNAP Facility")
			If IsDate(faci_entry_date) = True Then Call write_variable_in_case_note ("      - Facility Entry Date: " & faci_entry_date)
			If IsDate(faci_release_date) = True Then Call write_variable_in_case_note ("      - Release Date: " & faci_release_date)
			If release_date_unknown_checkbox = checked Then Call write_variable_in_case_note ("      - Release date is not known but is expected to be before " & day_30_from_application & ".")
			Call write_variable_in_case_note ("      - ")

		ElseIf deny_snap_due_to_faci = True Then
			Call write_variable_in_case_note ("* Resident is in a facility ")
			Call write_variable_in_case_note ("*   SNAP must be denied based on the current information.")
			If facility_name <> "" Then Call write_variable_in_case_note ("      - Facility Name: " & facility_name & " - an Ineligible SNAP Facility")
			If facility_name = "" Then Call write_variable_in_case_note ("      - Resident is in an Ineligible SNAP Facility")
			If IsDate(faci_entry_date) = True Then Call write_variable_in_case_note ("      - Facility Entry Date: " & faci_entry_date)
			If IsDate(faci_release_date) = True Then Call write_variable_in_case_note ("      - Release Date: " & faci_release_date)
			If release_date_unknown_checkbox = checked Then Call write_variable_in_case_note ("      - Release date is not known but is expected to be after " & day_30_from_application & ".")

		End If
		' household_in_a_facility_detail(delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, delay_explanation, snap_denial_explain, snap_denial_date, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
	End If
End If

Call write_variable_in_case_note ("---")

Call write_variable_in_case_note(worker_signature)

end_msg = "You have completed the EXPEDITED DETERMINATION." & vbCr & "Determination: " &  case_assesment_text
script_end_procedure_with_error_report (end_msg)
