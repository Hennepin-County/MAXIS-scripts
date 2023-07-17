'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CAF.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 1200                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

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
Call changelog_update("07/11/2023", "The CAF script will no longer allow for the selection of programs, MAXIS program status will inform which programs are included in the review of the case. ##~## ##~##It is vital that MAXIS is updated entirely before using the CAF script. In particular, the script will not operate if the CAF/Form Dates and Interview Dates have not been entered in the correct fields of MAXIS. ##~## ##~##It is best practice to run the CAF script PRIOR to any 'APP' as MAXIS updates are still possible and the display during the script run can support a secondary review of case details.##~##", "Casey Love, Hennepin County")
Call changelog_update("04/12/2023", "BUG Update - The Expedited Determination functionality was pulling Gross BUSI income, and we have updated it to pull NET income for the Expedited Determination information.", "Casey Love, Hennepin County")
call changelog_update("02/27/2023", "Reference updated for information about EBT cards. The button to open a webpage about EBT cards has been changed to open the current page mmanaged by Accounting instead of the previous Temporary Program Changes page.", "Casey Love, Hennepin County")
call changelog_update("12/22/2022", "BUG FIX - the NOTES - CAF script default income calculation for Expedited Determination was missing some UNEA, though the correct calculation could always be entered, the default count should include all income sources. Update to the script to ensure eash UNEA panel is added to the toatl before the first Expedited Determination information is shown.", "Casey Love, Hennepin County")
call changelog_update("11/14/2022", "Removed all handling for Interviews.##~## ##~##This script is not built to support the details of an interview or the documentation requirements. This script requires that an interview date has already been entered or that a CASE/NOTE has been created with NOTES - Interview. This script will end if interview date cannot be found on a case that an interview is required to process a CAF.", "Casey Love, Hennepin County")
call changelog_update("10/18/2022", "Removed Health Care renewal supports during the PHE. Health Care renewals remain paused.", "Ilse Ferris, Hennepin County")
call changelog_update("04/01/2022", "The functionality for Waiving an Interview has been removed. We can no longer waive SNAP Recertification Interviews.", "Casey Love, Hennepin County")
call changelog_update("03/17/2022", "BUG FIX##~##There have been reports of some of the required CASE:NOTEs missing from the script run at the end. We have updated some of the background functionality in this script to keep this from happening. If you notice issues with CASE:NOTEs missing at the end of this script run, report them to the BlueZone Script Team.", "Casey Love, Hennepin County")
call changelog_update("12/20/2021", "Removal of interview completed button. ", "MiKayla Handley")
call changelog_update("10/04/2021", "The CAF script now has 'SAVE YOUR WORK' functionality in testing. If the script fails, run it again on the same case and information should be filled back in to the dialog.##~##", "Casey Love, Hennepin County")
Call changelog_update("09/01/2021", "Expedited Determination Functionality has been completely enhanced.##~####~##The functionality to guide through the assesment of a case meeting expedited criteria has been updated. This new functionality adds a series of 3 new dialogs to support this process.##~####~##This new functionality matches the scripts NOTES - Expedited Determination and the new script NOTES - Interview.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/6/2020", "The script will now check the interview date entered on PROG or REVW to confirm the updated happened accurately when the script is tasked with updating PROG or REVW.##~## ##~##There may be a message that the update failed, this means this update must be completed manually.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/6/2020", "Added new functionality to have the script update the 'REVW' panel with the Interview Date and CAF date for Recertification Cases.##~## As this is a new functionality, please let us know how it works for you!.##~##", "Casey Love, Hennepin County")
Call changelog_update("10/09/2020", "Updated the functionality to adjust review dates for some cases to not require an interview date for Adult Cash Recertification cases.##~##", "Casey Love, Hennepin County")
Call changelog_update("10/01/2020", "Updated Standard Utility Allowances for 10/2020.", "Ilse Ferris, Hennepin County")
call changelog_update("09/01/2020", "UPDATE IN TESTING##~## ##~##There is new functionality added to this script to assist in updating REVW for some cases. There is detail in SIR about cases that need to have the next ER date adjusted due to a previous postponement. ##~## ##~##We have not had time with this functionality to complete live testing so all script runs will be a part of the test. Please let us know if you have any issues running this script.", "Casey Love, Hennepin County")
call changelog_update("05/18/2020", "Additional handling to be sure when saving the interview date to PROG the script does not get stuck.", "Casey Love, Hennepin County")
call changelog_update("04/02/2020", "BUG FIX - The 'notes on child support' detail was not always pulling into the case note. Updated to ensure the information entered in this field will be entered in the NOTE.##~##", "Casey Love, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality.", "Ilse Ferris")
Call changelog_update("01/29/2020", "When entering a expedited approval date or denial date on Dialog 8, the script will evaluate to be sure this date is not in the future. The script was requiring delay explanation when dates in the future were entered by accident, this additional handling will provide greater clarity of the updates needed.", "Casey Love, Hennepin County")
Call changelog_update("01/08/2020", "BUG FIX - When selecting CASH and SNAP for an MFIP Recertification, the script would error out and could not continue due to not being able to find the SNAP ER date on REVW. Updated the script to ignore that blank recert date.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/27/2019", "Added handling to support 10 or more children on STAT/PARE panels.", "Ilse Ferris, Hennepin County")
Call changelog_update("11/22/2019", "Added a checkbox on the verifications dialog pop-up. This checkbox will add detail to the verifications case note that there are verifications that have been postponed.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/22/2019", "Added handling for ID information and ID requirements for household members and AREP (if interviewed). This information is added to Dialog One.##~##This functionality mandates detail if the ID verification is 'Other' and is required.##~##", "Casey Love, Hennepin County")
Call changelog_update("11/14/2019", "BUG FIX - Dialog 4 had some fields overlapping each other sometimes, which made it difficult to read/update. Fixed the layout of Dialog  4 (CSES).##~##", "Casey Love, Hennepin County")
Call changelog_update("11/08/2019", "Added handling for the script to change a 4 digit footer year to a 2 digit footer year (2019 becomes 19) when entering recertification month and year by program. ##~##", "Casey Love, Hennepin County")
Call changelog_update("11/07/2019", "BUG FIX - Dialog 4 was sometimes too short. If there are a number of people with child support income, not all of the child support detail would be viewable as it would be taller than the computer screen. Updated the script so that Dialog 4 now has tabs if there are more than four members with child support income, so there are multiple pages of Dialog 4 (like Dialog 2 and 3).##~##", "Casey Love, Hennepin County")
Call changelog_update("10/16/2019", "BUG Fix - sometimes the script hit an error after leaving Dialog 8 - this should resolve that error. ##~## ##~## Added a NEW BUTTON that will display the Missing Fields Message (also called the 'Error Message') after clicking 'Done' on dialog 8 if the script needs updates. Look for the button 'Show Dialog Review Message' on each dialog after the message shows for the first time. ##~## This button will allow you to review the missing fields or updates that need to be made so that you do not have to try to remember them. The button only appears after the message was shown for the first time.##~##", "Casey Love, Hennepin County")
Call changelog_update("10/14/2019", "Added autofill functionality for TIME and SANC panels so the editboxes are filled if the panel is present.##~##", "Casey Love, Hennepin County")
call changelog_update("10/10/2019", "Updated 3 bugs/issues: ##~## ##~## - Sometimmes the list of clients on the 'Qualifying Quesitons Dialog' was not filled and was blank, this is now resolved and should always have a list of clients. ##~## - The script was 'forgetting' informmation typed into a ComboBox when a dialog appears for a subsequent time. This is now resolved. ##~## - Added headers to the mmissed fields/error message after Dialog 8 for more readability.", "Casey Love, Hennepin County")
Call changelog_update("10/01/2019", "CAF Functionality is enhanced for more complete and comprehensive documentation of CAF processing. This new functionality has been available for trial for the past 2 weeks. ##~## ##~## Live Skype Demos of this new functionality are availble this week and next. See Hot Topics for more details about the enhanceed functionality and the demo sessions. ##~##", "Casey Love, Hennepin County")
Call changelog_update("10/01/2019", "This script will be updated at the end of the day (10/1/2019) to the new CAF functionality. Additional details and resources can be found in Hot Topics or the BlueZone Script Team Sharepoint page.", "Casey Love, Hennepin County")
Call changelog_update("04/10/2019", "There was a bug that sometimes made the dialogs write over each other and be illegible, updated the script to keep this from happening.", "Casey Love, Hennepin County")
Call changelog_update("03/06/2019", "Added 2 new options to the Notes on Income button to support referencing CASE/NOTE made by Earned Income Budgeting.", "Casey Love, Hennepin County")
CALL changelog_update("10/17/2018", "Updated dialog box to reflect currecnt application process.", "MiKayla Handley, Hennepin County")
call changelog_update("05/05/2018", "Added autofill functionality for the DIET panel. NAT errors have been resolved.", "Ilse Ferris, Hennepin County")
call changelog_update("05/04/2018", "Removed autofill functionality for the DIET panel temporarily until MAXIS help desk can resolve NAT errors.", "Ilse Ferris, Hennepin County")
call changelog_update("01/11/2017", "Adding functionality to offer a TIKL for 12 month contact on 24 month SNAP renewals.", "Charles Potter, DHS")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THESE ARE EXP DET FUNXTION'
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
			case_assesment_text = "Case appears EXPEDITED and ready to approve"
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
			case_assesment_text = "Case appears EXPEDITED but approval must be delayed."
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
		case_assesment_text = "Case does NOT appear EXPEDITED, approval decision should follow standard SNAP Policy."
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

function save_your_work()
'This function records the variables into a txt file so that it can be retrieved by the script if run later.

	'Now determines name of file
	If MAXIS_case_number <> "" Then
		local_changelog_path = user_myDocs_folder & "caf-variables-" & MAXIS_case_number & "-info.txt"
	End If

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

			'Write the contents of the text file
            objTextStream.WriteLine "MAXIS_footer_month" & "^~^~^~^~^~^~^" & MAXIS_footer_month
            objTextStream.WriteLine "MAXIS_footer_year" & "^~^~^~^~^~^~^" & MAXIS_footer_year
            objTextStream.WriteLine "CAF_form" & "^~^~^~^~^~^~^" & CAF_form
            If number_verifs_checkbox = checked Then objTextStream.WriteLine "number_verifs_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verifs_postponed_checkbox = checked Then objTextStream.WriteLine "verifs_postponed_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"

            objTextStream.WriteLine "adult_cash_count" & "^~^~^~^~^~^~^" & adult_cash_count
            objTextStream.WriteLine "child_cash_count" & "^~^~^~^~^~^~^" & child_cash_count
            If pregnant_caregiver_checkbox = checked then objTextStream.WriteLine "pregnant_caregiver_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "adult_snap_count" & "^~^~^~^~^~^~^" & adult_snap_count
            objTextStream.WriteLine "child_snap_count" & "^~^~^~^~^~^~^" & child_snap_count
            objTextStream.WriteLine "adult_emer_count" & "^~^~^~^~^~^~^" & adult_emer_count
            objTextStream.WriteLine "child_emer_count" & "^~^~^~^~^~^~^" & child_emer_count
            objTextStream.WriteLine "EATS" & "^~^~^~^~^~^~^" & EATS
            objTextStream.WriteLine "relationship_detail" & "^~^~^~^~^~^~^" & relationship_detail

            objTextStream.WriteLine "determined_income" & "^~^~^~^~^~^~^" & determined_income
            objTextStream.WriteLine "income_review_completed" & "^~^~^~^~^~^~^" & income_review_completed
            objTextStream.WriteLine "jobs_income_yn" & "^~^~^~^~^~^~^" & jobs_income_yn
            objTextStream.WriteLine "busi_income_yn" & "^~^~^~^~^~^~^" & busi_income_yn
            objTextStream.WriteLine "unea_income_yn" & "^~^~^~^~^~^~^" & unea_income_yn
            ' For exp_jobs = 0 to UBound(JOBS_ARRAY, 2)
            '     objTextStream.WriteLine "JOBS_ARRAY" & "^~^~^~^~^~^~^" &
            ' Next
            ' For exp_busi = 0 to UBound(BUSI_ARRAY, 2)
            '     objTextStream.WriteLine "BUSI_ARRAY" & "^~^~^~^~^~^~^" &
            ' Next
            ' For exp_unea = 0 to UBound(UNEA_ARRAY, 2)
            '     objTextStream.WriteLine "UNEA_ARRAY" & "^~^~^~^~^~^~^" &
            ' Next

            objTextStream.WriteLine "determined_assets" & "^~^~^~^~^~^~^" & determined_assets
            objTextStream.WriteLine "assets_review_completed" & "^~^~^~^~^~^~^" & assets_review_completed
            objTextStream.WriteLine "cash_amount_yn" & "^~^~^~^~^~^~^" & cash_amount_yn
            objTextStream.WriteLine "bank_account_yn" & "^~^~^~^~^~^~^" & bank_account_yn
            objTextStream.WriteLine "cash_amount" & "^~^~^~^~^~^~^" & cash_amount
            ' For exp_acct = 0 to UBound(ACCOUNTS_ARRAY, 2)
            '     objTextStream.WriteLine "ACCOUNTS_ARRAY" & "^~^~^~^~^~^~^" &
            ' Next

            objTextStream.WriteLine "determined_shel" & "^~^~^~^~^~^~^" & determined_shel
            objTextStream.WriteLine "shel_review_completed" & "^~^~^~^~^~^~^" & shel_review_completed
            objTextStream.WriteLine "rent_amount" & "^~^~^~^~^~^~^" & rent_amount
            objTextStream.WriteLine "lot_rent_amount" & "^~^~^~^~^~^~^" & lot_rent_amount
            objTextStream.WriteLine "mortgage_amount" & "^~^~^~^~^~^~^" & mortgage_amount
            objTextStream.WriteLine "insurance_amount" & "^~^~^~^~^~^~^" & insurance_amount
            objTextStream.WriteLine "tax_amount" & "^~^~^~^~^~^~^" & tax_amount
            objTextStream.WriteLine "room_amount" & "^~^~^~^~^~^~^" & room_amount
            objTextStream.WriteLine "garage_amount" & "^~^~^~^~^~^~^" & garage_amount
            objTextStream.WriteLine "subsidy_amount" & "^~^~^~^~^~^~^" & subsidy_amount

            objTextStream.WriteLine "determined_utilities" & "^~^~^~^~^~^~^" & determined_utilities
            objTextStream.WriteLine "heat_expense" & "^~^~^~^~^~^~^" & heat_expense
            objTextStream.WriteLine "ac_expense" & "^~^~^~^~^~^~^" & ac_expense
            objTextStream.WriteLine "electric_expense" & "^~^~^~^~^~^~^" & electric_expense
            objTextStream.WriteLine "phone_expense" & "^~^~^~^~^~^~^" & phone_expense
            objTextStream.WriteLine "none_expense" & "^~^~^~^~^~^~^" & none_expense
            objTextStream.WriteLine "all_utilities" & "^~^~^~^~^~^~^" & all_utilities

            objTextStream.WriteLine "do_we_have_applicant_id" & "^~^~^~^~^~^~^" & do_we_have_applicant_id

            If CASH_checkbox = checked then objTextStream.WriteLine "CASH_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If GRH_checkbox = checked then objTextStream.WriteLine "GRH_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If SNAP_checkbox = checked then objTextStream.WriteLine "SNAP_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If HC_checkbox = checked then objTextStream.WriteLine "HC_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If EMER_checkbox = checked then objTextStream.WriteLine "EMER_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"

            objTextStream.WriteLine "multiple_CAF_dates" & "^~^~^~^~^~^~^" & multiple_CAF_dates
            objTextStream.WriteLine "multiple_interview_dates" & "^~^~^~^~^~^~^" & multiple_interview_dates
            objTextStream.WriteLine "adult_cash" & "^~^~^~^~^~^~^" & adult_cash
            objTextStream.WriteLine "family_cash" & "^~^~^~^~^~^~^" & family_cash
            objTextStream.WriteLine "the_process_for_cash" & "^~^~^~^~^~^~^" & the_process_for_cash
            objTextStream.WriteLine "type_of_cash" & "^~^~^~^~^~^~^" & type_of_cash
            objTextStream.WriteLine "cash_recert_mo" & "^~^~^~^~^~^~^" & cash_recert_mo
            objTextStream.WriteLine "cash_recert_yr" & "^~^~^~^~^~^~^" & cash_recert_yr
            objTextStream.WriteLine "the_process_for_grh" & "^~^~^~^~^~^~^" & the_process_for_grh
            objTextStream.WriteLine "grh_recert_mo" & "^~^~^~^~^~^~^" & grh_recert_mo
            objTextStream.WriteLine "grh_recert_yr" & "^~^~^~^~^~^~^" & grh_recert_yr
            objTextStream.WriteLine "the_process_for_snap" & "^~^~^~^~^~^~^" & the_process_for_snap
            objTextStream.WriteLine "snap_recert_mo" & "^~^~^~^~^~^~^" & snap_recert_mo
            objTextStream.WriteLine "snap_recert_yr" & "^~^~^~^~^~^~^" & snap_recert_yr
            objTextStream.WriteLine "the_process_for_hc" & "^~^~^~^~^~^~^" & the_process_for_hc
            objTextStream.WriteLine "hc_recert_mo" & "^~^~^~^~^~^~^" & hc_recert_mo
            objTextStream.WriteLine "hc_recert_yr" & "^~^~^~^~^~^~^" & hc_recert_yr
            objTextStream.WriteLine "the_process_for_emer" & "^~^~^~^~^~^~^" & the_process_for_emer
            objTextStream.WriteLine "type_of_emer" & "^~^~^~^~^~^~^" & type_of_emer
            ' objTextStream.WriteLine "CAF_type" & "^~^~^~^~^~^~^" & CAF_type
			objTextStream.WriteLine "application_processing" & "^~^~^~^~^~^~^" & application_processing
			objTextStream.WriteLine "recert_processing" & "^~^~^~^~^~^~^" & recert_processing
            objTextStream.WriteLine "CAF_datestamp" & "^~^~^~^~^~^~^" & CAF_datestamp
            objTextStream.WriteLine "PROG_CAF_datestamp" & "^~^~^~^~^~^~^" & PROG_CAF_datestamp
            objTextStream.WriteLine "REVW_CAF_datestamp" & "^~^~^~^~^~^~^" & REVW_CAF_datestamp
            objTextStream.WriteLine "interview_date" & "^~^~^~^~^~^~^" & interview_date
            objTextStream.WriteLine "PROG_interview_date" & "^~^~^~^~^~^~^" & PROG_interview_date
            objTextStream.WriteLine "REVW_interview_date" & "^~^~^~^~^~^~^" & REVW_interview_date
            objTextStream.WriteLine "case_details_and_notes_about_process" & "^~^~^~^~^~^~^" & case_details_and_notes_about_process
            objTextStream.WriteLine "SNAP_recert_is_likely_24_months" & "^~^~^~^~^~^~^" & SNAP_recert_is_likely_24_months
            objTextStream.WriteLine "exp_screening_note_found" & "^~^~^~^~^~^~^" & exp_screening_note_found
            objTextStream.WriteLine "interview_required" & "^~^~^~^~^~^~^" & interview_required
            objTextStream.WriteLine "xfs_screening" & "^~^~^~^~^~^~^" & xfs_screening
            objTextStream.WriteLine "xfs_screening_display" & "^~^~^~^~^~^~^" & xfs_screening_display
            objTextStream.WriteLine "caf_one_income" & "^~^~^~^~^~^~^" & caf_one_income
            objTextStream.WriteLine "caf_one_assets" & "^~^~^~^~^~^~^" & caf_one_assets
            objTextStream.WriteLine "caf_one_resources" & "^~^~^~^~^~^~^" & caf_one_resources
            objTextStream.WriteLine "caf_one_rent" & "^~^~^~^~^~^~^" & caf_one_rent
            objTextStream.WriteLine "caf_one_utilities" & "^~^~^~^~^~^~^" & caf_one_utilities
            objTextStream.WriteLine "caf_one_expenses" & "^~^~^~^~^~^~^" & caf_one_expenses
            objTextStream.WriteLine "exp_det_case_note_found" & "^~^~^~^~^~^~^" & exp_det_case_note_found
            objTextStream.WriteLine "snap_exp_yn" & "^~^~^~^~^~^~^" & snap_exp_yn
            objTextStream.WriteLine "snap_denial_date" & "^~^~^~^~^~^~^" & snap_denial_date
            objTextStream.WriteLine "interview_completed_case_note_found" & "^~^~^~^~^~^~^" & interview_completed_case_note_found
            objTextStream.WriteLine "interview_with" & "^~^~^~^~^~^~^" & interview_with
            objTextStream.WriteLine "interview_type" & "^~^~^~^~^~^~^" & interview_type
            objTextStream.WriteLine "verifications_requested_case_note_found" & "^~^~^~^~^~^~^" & verifications_requested_case_note_found
            objTextStream.WriteLine "verifs_needed" & "^~^~^~^~^~^~^" & verifs_needed
            If verif_snap_checkbox = checked then objTextStream.WriteLine "verif_snap_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_cash_checkbox = checked then objTextStream.WriteLine "verif_cash_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_mfip_checkbox = checked then objTextStream.WriteLine "verif_mfip_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_dwp_checkbox = checked then objTextStream.WriteLine "verif_dwp_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_msa_checkbox = checked then objTextStream.WriteLine "verif_msa_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_ga_checkbox = checked then objTextStream.WriteLine "verif_ga_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_grh_checkbox = checked then objTextStream.WriteLine "verif_grh_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_emer_checkbox = checked then objTextStream.WriteLine "verif_emer_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If verif_hc_checkbox = checked then objTextStream.WriteLine "verif_hc_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"

            objTextStream.WriteLine "caf_qualifying_questions_case_note_found" & "^~^~^~^~^~^~^" & caf_qualifying_questions_case_note_found
            objTextStream.WriteLine "qual_question_one" & "^~^~^~^~^~^~^" & qual_question_one
            objTextStream.WriteLine "qual_memb_one" & "^~^~^~^~^~^~^" & qual_memb_one
            objTextStream.WriteLine "qual_question_two" & "^~^~^~^~^~^~^" & qual_question_two
            objTextStream.WriteLine "qual_memb_two" & "^~^~^~^~^~^~^" & qual_memb_two
            objTextStream.WriteLine "qual_question_three" & "^~^~^~^~^~^~^" & qual_question_three
            objTextStream.WriteLine "qual_memb_three" & "^~^~^~^~^~^~^" & qual_memb_three
            objTextStream.WriteLine "qual_question_four" & "^~^~^~^~^~^~^" & qual_question_four
            objTextStream.WriteLine "qual_memb_four" & "^~^~^~^~^~^~^" & qual_memb_four
            objTextStream.WriteLine "qual_question_five" & "^~^~^~^~^~^~^" & qual_question_five
            objTextStream.WriteLine "qual_memb_five" & "^~^~^~^~^~^~^" & qual_memb_five
            objTextStream.WriteLine "appt_notc_sent_on" & "^~^~^~^~^~^~^" & appt_notc_sent_on
            objTextStream.WriteLine "appt_date_in_note" & "^~^~^~^~^~^~^" & appt_date_in_note
            For each HH_MEMB in HH_member_array
                objTextStream.WriteLine "HH_member_array" & "^~^~^~^~^~^~^" & HH_MEMB
            Next
            objTextStream.WriteLine "addr_line_one" & "^~^~^~^~^~^~^" & addr_line_one
            objTextStream.WriteLine "addr_line_two" & "^~^~^~^~^~^~^" & addr_line_two
            objTextStream.WriteLine "city" & "^~^~^~^~^~^~^" & city
            objTextStream.WriteLine "state" & "^~^~^~^~^~^~^" & state
            objTextStream.WriteLine "zip" & "^~^~^~^~^~^~^" & zip
            objTextStream.WriteLine "addr_county" & "^~^~^~^~^~^~^" & addr_county
            objTextStream.WriteLine "homeless_yn" & "^~^~^~^~^~^~^" & homeless_yn
            objTextStream.WriteLine "reservation_yn" & "^~^~^~^~^~^~^" & reservation_yn
            objTextStream.WriteLine "addr_verif" & "^~^~^~^~^~^~^" & addr_verif
            objTextStream.WriteLine "living_situation" & "^~^~^~^~^~^~^" & living_situation
            objTextStream.WriteLine "addr_eff_date" & "^~^~^~^~^~^~^" & addr_eff_date
            objTextStream.WriteLine "addr_future_date" & "^~^~^~^~^~^~^" & addr_future_date
            objTextStream.WriteLine "mail_line_one" & "^~^~^~^~^~^~^" & mail_line_one
            objTextStream.WriteLine "mail_line_two" & "^~^~^~^~^~^~^" & mail_line_two
            objTextStream.WriteLine "mail_city_line" & "^~^~^~^~^~^~^" & mail_city_line
            objTextStream.WriteLine "mail_state_line" & "^~^~^~^~^~^~^" & mail_state_line
            objTextStream.WriteLine "mail_zip_line" & "^~^~^~^~^~^~^" & mail_zip_line
            objTextStream.WriteLine "notes_on_address" & "^~^~^~^~^~^~^" & notes_on_address
            For the_members = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                box_one_info = ""
                box_two_info = ""
                box_three_info = ""
                box_four_info = ""
                box_five_info = ""
                box_six_info = ""
                box_seven_info = ""
                box_eight_info = ""
                box_nine_info = ""
                If ALL_MEMBERS_ARRAY(include_cash_checkbox, the_members) = checked Then box_one_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(include_snap_checkbox, the_members) = checked Then box_two_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(include_emer_checkbox, the_members) = checked Then box_three_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(count_cash_checkbox, the_members) = checked Then box_four_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(count_snap_checkbox, the_members) = checked Then box_five_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(count_emer_checkbox, the_members) = checked Then box_six_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(pwe_checkbox, the_members) = checked Then box_seven_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(shel_verif_checkbox, the_members) = checked Then box_eight_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(id_required, the_members) = checked Then box_nine_info = "CHECKED"

                objTextStream.WriteLine "ALL_MEMBERS_ARRAY" & "^~^~^~^~^~^~^" &ALL_MEMBERS_ARRAY(memb_numb, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_name, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_age, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(full_clt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_id_verif, the_members)&"~**^"&box_one_info&"~**^"&box_two_info&"~**^"&box_three_info&"~**^"&box_four_info&"~**^"&box_five_info&"~**^"&box_six_info&"~**^"&_
                ALL_MEMBERS_ARRAY(clt_wreg_status, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_abawd_status, the_members)&"~**^"&box_seven_info&"~**^"&ALL_MEMBERS_ARRAY(numb_abawd_used, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(list_abawd_mo, the_members)&"~**^"&ALL_MEMBERS_ARRAY(first_second_set, the_members)&"~**^"&ALL_MEMBERS_ARRAY(list_second_set, the_members)&"~**^"&ALL_MEMBERS_ARRAY(explain_no_second, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(numb_banked_mo, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_abawd_notes, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_exists, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_subsudized, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_shared, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_rent_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_rent_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_lot_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_lot_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_ins_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_ins_verif,the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_tax_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_tax_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_room_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_room_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_room_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_room_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_garage_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_garage_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt,the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(wreg_exists, the_members)&"~**^"&box_eight_info&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_verif_added, the_members)&"~**^"&ALL_MEMBERS_ARRAY(gather_detail, the_members)&"~**^"&ALL_MEMBERS_ARRAY(id_detail, the_members)&"~**^"&box_nine_info&"~**^"&_
                ALL_MEMBERS_ARRAY(clt_notes, the_members)
            Next
            objTextStream.WriteLine "total_shelter_amount" & "^~^~^~^~^~^~^" & total_shelter_amount
            objTextStream.WriteLine "full_shelter_details" & "^~^~^~^~^~^~^" & full_shelter_details
            objTextStream.WriteLine "shelter_details" & "^~^~^~^~^~^~^" & shelter_details
            objTextStream.WriteLine "shelter_details_two" & "^~^~^~^~^~^~^" & shelter_details_two
            objTextStream.WriteLine "shelter_details_three" & "^~^~^~^~^~^~^" & shelter_details_three
            objTextStream.WriteLine "prosp_heat_air" & "^~^~^~^~^~^~^" & prosp_heat_air
            objTextStream.WriteLine "prosp_electric" & "^~^~^~^~^~^~^" & prosp_electric
            objTextStream.WriteLine "prosp_phone" & "^~^~^~^~^~^~^" & prosp_phone
            objTextStream.WriteLine "hest_information" & "^~^~^~^~^~^~^" & hest_information
            objTextStream.WriteLine "ABPS" & "^~^~^~^~^~^~^" & ABPS
            objTextStream.WriteLine "ACCI" & "^~^~^~^~^~^~^" & ACCI
            objTextStream.WriteLine "notes_on_acct" & "^~^~^~^~^~^~^" & notes_on_acct
            objTextStream.WriteLine "notes_on_acut" & "^~^~^~^~^~^~^" & notes_on_acut
            objTextStream.WriteLine "AREP" & "^~^~^~^~^~^~^" & AREP
            objTextStream.WriteLine "BILS" & "^~^~^~^~^~^~^" & BILS
            objTextStream.WriteLine "notes_on_cash" & "^~^~^~^~^~^~^" & notes_on_cash
            objTextStream.WriteLine "notes_on_cars" & "^~^~^~^~^~^~^" & notes_on_cars
            objTextStream.WriteLine "notes_on_coex" & "^~^~^~^~^~^~^" & notes_on_coex
            objTextStream.WriteLine "notes_on_dcex" & "^~^~^~^~^~^~^" & notes_on_dcex
            objTextStream.WriteLine "DIET" & "^~^~^~^~^~^~^" & DIET
            objTextStream.WriteLine "DISA" & "^~^~^~^~^~^~^" & DISA
            objTextStream.WriteLine "EMPS" & "^~^~^~^~^~^~^" & EMPS
            objTextStream.WriteLine "FACI" & "^~^~^~^~^~^~^" & FACI
            objTextStream.WriteLine "FMED" & "^~^~^~^~^~^~^" & FMED
            objTextStream.WriteLine "IMIG" & "^~^~^~^~^~^~^" & IMIG
            objTextStream.WriteLine "INSA" & "^~^~^~^~^~^~^" & INSA
            For the_jobs = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                box_one_info = ""
                box_two_info = ""
                If ALL_JOBS_PANELS_ARRAY(estimate_only, the_jobs) = checked Then box_one_info = "CHECKED"
                If ALL_JOBS_PANELS_ARRAY(verif_checkbox, the_jobs) = checked Then box_two_info = "CHECKED"
                objTextStream.WriteLine "ALL_JOBS_PANELS_ARRAY" & "^~^~^~^~^~^~^" &ALL_JOBS_PANELS_ARRAY(memb_numb, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(panel_instance, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(employer_name, the_jobs)&"~**^"&box_one_info&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_explain, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_code, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(info_month, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(hrly_wage, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(main_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(job_retro_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(job_prosp_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(retro_hours, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(prosp_hours, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_prosp_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_calc_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(EI_case_note, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_calc_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_prosp_income, the_jobs)&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&ALL_JOBS_PANELS_ARRAY(start_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(end_date, the_jobs)&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&box_two_info&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_added, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(budget_explain, the_jobs)
            Next
            For the_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                box_one_info = ""
                box_two_info = ""
                box_three_info = ""
                If ALL_BUSI_PANELS_ARRAY(estimate_only, the_busi) = checked Then box_one_info = "CHECKED"
                If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, the_busi) = checked Then box_two_info = "CHECKED"
                If ALL_BUSI_PANELS_ARRAY(verif_checkbox, the_busi) = checked Then box_three_info = "CHECKED"

                objTextStream.WriteLine "ALL_BUSI_PANELS_ARRAY" & "^~^~^~^~^~^~^" &ALL_BUSI_PANELS_ARRAY(memb_numb, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(panel_instance, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_type, the_busi)&"~**^"&box_one_info&"~**^"&ALL_BUSI_PANELS_ARRAY(verif_explain, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(calc_method, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(info_month, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(mthd_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_ret_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_pro_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(cash_income_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_ret_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_pro_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(cash_expense_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_ret_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_pro_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(snap_income_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_ret_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_pro_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(snap_expense_verif, the_busi)&"~**^"&box_two_info&"~**^"&ALL_BUSI_PANELS_ARRAY(start_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(end_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_desc, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_structure, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(share_num, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(share_denom, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(partners_in_HH, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(exp_not_allwd, the_busi)&"~**^"&box_three_info&"~**^"&ALL_BUSI_PANELS_ARRAY(verif_added, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(budget_explain, the_busi)
            Next
            objTextStream.WriteLine "cit_id" & "^~^~^~^~^~^~^" & cit_id
            objTextStream.WriteLine "other_assets" & "^~^~^~^~^~^~^" & other_assets
            objTextStream.WriteLine "case_changes" & "^~^~^~^~^~^~^" & case_changes
            objTextStream.WriteLine "PREG" & "^~^~^~^~^~^~^" & PREG
            objTextStream.WriteLine "earned_income" & "^~^~^~^~^~^~^" & earned_income
            objTextStream.WriteLine "notes_on_rest" & "^~^~^~^~^~^~^" & notes_on_rest
            objTextStream.WriteLine "SCHL" & "^~^~^~^~^~^~^" & SCHL
            objTextStream.WriteLine "notes_on_jobs" & "^~^~^~^~^~^~^" & notes_on_jobs
            objTextStream.WriteLine "notes_on_cses" & "^~^~^~^~^~^~^" & notes_on_cses
            objTextStream.WriteLine "notes_on_time" & "^~^~^~^~^~^~^" & notes_on_time
            objTextStream.WriteLine "notes_on_sanction" & "^~^~^~^~^~^~^" & notes_on_sanction
            For the_unea = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                objTextStream.WriteLine "UNEA_INCOME_ARRAY" & "^~^~^~^~^~^~^" &UNEA_INCOME_ARRAY(memb_numb, the_unea)&"~**^"&UNEA_INCOME_ARRAY(panel_instance, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_type, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_month, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_verif, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_prosp_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_retro_amt, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_SNAP_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_pay_freq, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_pic_date_calc, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_start_date, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_notes, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_account_balance, the_unea)&"~**^"&UNEA_INCOME_ARRAY(direct_CS_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_amt, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(disb_CS_arrears_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(direct_CS_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_notes, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(disb_CS_months, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_prosp_budg, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_months, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_budg, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_RSDI_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_RSDI_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SSI_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SSI_notes, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UC_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(CS_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(SSA_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(calc_button, the_unea)&"~**^"&UNEA_INCOME_ARRAY(budget_notes, the_unea)
            Next
            objTextStream.WriteLine "notes_on_wreg" & "^~^~^~^~^~^~^" & notes_on_wreg
            objTextStream.WriteLine "full_abawd_info" & "^~^~^~^~^~^~^" & full_abawd_info
            objTextStream.WriteLine "notes_on_abawd" & "^~^~^~^~^~^~^" & notes_on_abawd
            objTextStream.WriteLine "notes_on_abawd_two" & "^~^~^~^~^~^~^" & notes_on_abawd_two
            objTextStream.WriteLine "notes_on_abawd_three" & "^~^~^~^~^~^~^" & notes_on_abawd_three
            objTextStream.WriteLine "programs_applied_for" & "^~^~^~^~^~^~^" & programs_applied_for
            If TIKL_checkbox = checked Then objTextStream.WriteLine "TIKL_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "interview_memb_list" & "^~^~^~^~^~^~^" & interview_memb_list
            objTextStream.WriteLine "shel_memb_list" & "^~^~^~^~^~^~^" & shel_memb_list
            objTextStream.WriteLine "verification_memb_list" & "^~^~^~^~^~^~^" & verification_memb_list
            objTextStream.WriteLine "notes_on_busi" & "^~^~^~^~^~^~^" & notes_on_busi
            'DLG 1
            If Used_Interpreter_checkbox = checked Then objTextStream.WriteLine "Used_Interpreter_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            ' objTextStream.WriteLine "how_app_rcvd" & "^~^~^~^~^~^~^" & how_app_rcvd
            objTextStream.WriteLine "arep_id_info" & "^~^~^~^~^~^~^" & arep_id_info
            objTextStream.WriteLine "CS_forms_sent_date" & "^~^~^~^~^~^~^" & CS_forms_sent_date
            objTextStream.WriteLine "case_changes" & "^~^~^~^~^~^~^" & case_changes
            'DLG 5'
            objTextStream.WriteLine "notes_on_ssa_income" & "^~^~^~^~^~^~^" & notes_on_ssa_income
            objTextStream.WriteLine "notes_on_VA_income" & "^~^~^~^~^~^~^" & notes_on_VA_income
            objTextStream.WriteLine "notes_on_WC_income" & "^~^~^~^~^~^~^" & notes_on_WC_income
            objTextStream.WriteLine "other_uc_income_notes" & "^~^~^~^~^~^~^" & other_uc_income_notes
            objTextStream.WriteLine "notes_on_other_UNEA" & "^~^~^~^~^~^~^" & notes_on_other_UNEA

            objTextStream.WriteLine "hest_information" & "^~^~^~^~^~^~^" & hest_information
            objTextStream.WriteLine "notes_on_acut" & "^~^~^~^~^~^~^" & notes_on_acut
            objTextStream.WriteLine "notes_on_coex" & "^~^~^~^~^~^~^" & notes_on_coex
            objTextStream.WriteLine "notes_on_dcex" & "^~^~^~^~^~^~^" & notes_on_dcex
            objTextStream.WriteLine "notes_on_other_deduction" & "^~^~^~^~^~^~^" & notes_on_other_deduction
            objTextStream.WriteLine "expense_notes" & "^~^~^~^~^~^~^" & expense_notes
            If address_confirmation_checkbox = checked Then objTextStream.WriteLine "address_confirmation_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "manual_total_shelter" & "^~^~^~^~^~^~^" & manual_total_shelter
            objTextStream.WriteLine "manual_amount_used" & "^~^~^~^~^~^~^" & manual_amount_used
            objTextStream.WriteLine "app_month_assets" & "^~^~^~^~^~^~^" & app_month_assets
            If confirm_no_account_panel_checkbox = checked Then objTextStream.WriteLine "confirm_no_account_panel_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "notes_on_other_assets" & "^~^~^~^~^~^~^" & notes_on_other_assets
            objTextStream.WriteLine "MEDI" & "^~^~^~^~^~^~^" & MEDI
            objTextStream.WriteLine "DISQ" & "^~^~^~^~^~^~^" & DISQ
            'EXP DET'
            objTextStream.WriteLine "full_determination_done" & "^~^~^~^~^~^~^" & full_determination_done
            ' Call run_expedited_determination_script_functionality(
            objTextStream.WriteLine "xfs_screening" & "^~^~^~^~^~^~^" & xfs_screening
            objTextStream.WriteLine "caf_one_income" & "^~^~^~^~^~^~^" & caf_one_income
            objTextStream.WriteLine "caf_one_assets" & "^~^~^~^~^~^~^" & caf_one_assets
            objTextStream.WriteLine "caf_one_rent" & "^~^~^~^~^~^~^" & caf_one_rent
            objTextStream.WriteLine "caf_one_utilities" & "^~^~^~^~^~^~^" & caf_one_utilities
            objTextStream.WriteLine "determined_income" & "^~^~^~^~^~^~^" & determined_income
            objTextStream.WriteLine "determined_assets" & "^~^~^~^~^~^~^" & determined_assets
            objTextStream.WriteLine "determined_shel" & "^~^~^~^~^~^~^" & determined_shel
            objTextStream.WriteLine "determined_utilities" & "^~^~^~^~^~^~^" & determined_utilities
            objTextStream.WriteLine "calculated_resources" & "^~^~^~^~^~^~^" & calculated_resources
            objTextStream.WriteLine "calculated_expenses" & "^~^~^~^~^~^~^" & calculated_expenses
            objTextStream.WriteLine "calculated_low_income_asset_test" & "^~^~^~^~^~^~^" & calculated_low_income_asset_test
            objTextStream.WriteLine "calculated_resources_less_than_expenses_test" & "^~^~^~^~^~^~^" & calculated_resources_less_than_expenses_test
            objTextStream.WriteLine "is_elig_XFS" & "^~^~^~^~^~^~^" & is_elig_XFS
            objTextStream.WriteLine "approval_date" & "^~^~^~^~^~^~^" & approval_date
            objTextStream.WriteLine "applicant_id_on_file_yn" & "^~^~^~^~^~^~^" & applicant_id_on_file_yn
            objTextStream.WriteLine "applicant_id_through_SOLQ" & "^~^~^~^~^~^~^" & applicant_id_through_SOLQ
            objTextStream.WriteLine "delay_explanation" & "^~^~^~^~^~^~^" & delay_explanation
            objTextStream.WriteLine "snap_denial_date" & "^~^~^~^~^~^~^" & snap_denial_date
            objTextStream.WriteLine "snap_denial_explain" & "^~^~^~^~^~^~^" & snap_denial_explain
            objTextStream.WriteLine "case_assesment_text" & "^~^~^~^~^~^~^" & case_assesment_text
            objTextStream.WriteLine "next_steps_one" & "^~^~^~^~^~^~^" & next_steps_one
            objTextStream.WriteLine "next_steps_two" & "^~^~^~^~^~^~^" & next_steps_two
            objTextStream.WriteLine "next_steps_three" & "^~^~^~^~^~^~^" & next_steps_three
            objTextStream.WriteLine "next_steps_four" & "^~^~^~^~^~^~^" & next_steps_four
            objTextStream.WriteLine "postponed_verifs_yn" & "^~^~^~^~^~^~^" & postponed_verifs_yn
            objTextStream.WriteLine "list_postponed_verifs" & "^~^~^~^~^~^~^" & list_postponed_verifs
            objTextStream.WriteLine "day_30_from_application" & "^~^~^~^~^~^~^" & day_30_from_application
            objTextStream.WriteLine "other_snap_state" & "^~^~^~^~^~^~^" & other_snap_state
            objTextStream.WriteLine "other_state_reported_benefit_end_date" & "^~^~^~^~^~^~^" & other_state_reported_benefit_end_date
            objTextStream.WriteLine "other_state_benefits_openended" & "^~^~^~^~^~^~^" & other_state_benefits_openended
            objTextStream.WriteLine "other_state_contact_yn" & "^~^~^~^~^~^~^" & other_state_contact_yn
            objTextStream.WriteLine "other_state_verified_benefit_end_date" & "^~^~^~^~^~^~^" & other_state_verified_benefit_end_date
            objTextStream.WriteLine "mn_elig_begin_date" & "^~^~^~^~^~^~^" & mn_elig_begin_date
            objTextStream.WriteLine "action_due_to_out_of_state_benefits" & "^~^~^~^~^~^~^" & action_due_to_out_of_state_benefits
            objTextStream.WriteLine "case_has_previously_postponed_verifs_that_prevent_exp_snap" & "^~^~^~^~^~^~^" & case_has_previously_postponed_verifs_that_prevent_exp_snap
            objTextStream.WriteLine "prev_post_verif_assessment_done" & "^~^~^~^~^~^~^" & prev_post_verif_assessment_done
            objTextStream.WriteLine "previous_date_of_application" & "^~^~^~^~^~^~^" & previous_date_of_application
            objTextStream.WriteLine "previous_expedited_package" & "^~^~^~^~^~^~^" & previous_expedited_package
            objTextStream.WriteLine "prev_verifs_mandatory_yn" & "^~^~^~^~^~^~^" & prev_verifs_mandatory_yn
            objTextStream.WriteLine "prev_verif_list" & "^~^~^~^~^~^~^" & prev_verif_list
            objTextStream.WriteLine "curr_verifs_postponed_yn" & "^~^~^~^~^~^~^" & curr_verifs_postponed_yn
            objTextStream.WriteLine "ongoing_snap_approved_yn" & "^~^~^~^~^~^~^" & ongoing_snap_approved_yn
            objTextStream.WriteLine "prev_post_verifs_recvd_yn" & "^~^~^~^~^~^~^" & prev_post_verifs_recvd_yn
            objTextStream.WriteLine "delay_action_due_to_faci" & "^~^~^~^~^~^~^" & delay_action_due_to_faci
            objTextStream.WriteLine "deny_snap_due_to_faci" & "^~^~^~^~^~^~^" & deny_snap_due_to_faci
            objTextStream.WriteLine "faci_review_completed" & "^~^~^~^~^~^~^" & faci_review_completed
            objTextStream.WriteLine "facility_name" & "^~^~^~^~^~^~^" & facility_name
            objTextStream.WriteLine "snap_inelig_faci_yn" & "^~^~^~^~^~^~^" & snap_inelig_faci_yn
            objTextStream.WriteLine "faci_entry_date" & "^~^~^~^~^~^~^" & faci_entry_date
            objTextStream.WriteLine "faci_release_date" & "^~^~^~^~^~^~^" & faci_release_date
            If release_date_unknown_checkbox = checked Then objTextStream.WriteLine "release_date_unknown_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "release_within_30_days_yn" & "^~^~^~^~^~^~^" & release_within_30_days_yn

            objTextStream.WriteLine "next_er_month" & "^~^~^~^~^~^~^" & next_er_month
            objTextStream.WriteLine "next_er_year" & "^~^~^~^~^~^~^" & next_er_year
            objTextStream.WriteLine "CAF_status" & "^~^~^~^~^~^~^" & CAF_status
            objTextStream.WriteLine "actions_taken" & "^~^~^~^~^~^~^" & actions_taken
            If application_signed_checkbox = checked Then objTextStream.WriteLine "application_signed_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If eDRS_sent_checkbox = checked Then objTextStream.WriteLine "eDRS_sent_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If updated_MMIS_checkbox = checked Then objTextStream.WriteLine "updated_MMIS_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If WF1_checkbox = checked Then objTextStream.WriteLine "WF1_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If Sent_arep_checkbox = checked Then objTextStream.WriteLine "Sent_arep_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If intake_packet_checkbox = checked Then objTextStream.WriteLine "intake_packet_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If IAA_checkbox = checked Then objTextStream.WriteLine "IAA_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If recert_period_checkbox = checked Then objTextStream.WriteLine "recert_period_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If R_R_checkbox = checked Then objTextStream.WriteLine "R_R_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If E_and_T_checkbox = checked Then objTextStream.WriteLine "E_and_T_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If elig_req_explained_checkbox = checked Then objTextStream.WriteLine "elig_req_explained_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If benefit_payment_explained_checkbox = checked Then objTextStream.WriteLine "benefit_payment_explained_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "other_notes" & "^~^~^~^~^~^~^" & other_notes
            If client_delay_checkbox = checked Then objTextStream.WriteLine "client_delay_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If TIKL_checkbox = checked Then objTextStream.WriteLine "TIKL_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            If client_delay_TIKL_checkbox = checked Then objTextStream.WriteLine "client_delay_TIKL_checkbox" & "^~^~^~^~^~^~^" & "CHECKED"
            objTextStream.WriteLine "verif_req_form_sent_date" & "^~^~^~^~^~^~^" & verif_req_form_sent_date
            objTextStream.WriteLine "worker_signature" & "^~^~^~^~^~^~^" & worker_signature
            objTextStream.WriteLine "script_information_was_restored" & "^~^~^~^~^~^~^" & script_information_was_restored

            objTextStream.WriteLine case_notes_information
            ' objTextStream.WriteLine "" & "^~^~^~^~^~^~^" &

            'Close the object so it can be opened again shortly
			objTextStream.Close

            script_run_lowdown = ""

            script_run_lowdown = script_run_lowdown & vbCr & "MAXIS_footer_month" & ": " & MAXIS_footer_month
            script_run_lowdown = script_run_lowdown & vbCr & "MAXIS_footer_year" & ": " & MAXIS_footer_year
            script_run_lowdown = script_run_lowdown & vbCr & "CAF_form" & ": " & CAF_form
            If number_verifs_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "number_verifs_checkbox" & ": " & "CHECKED"
            If number_verifs_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "number_verifs_checkbox" & ": " & "UNCHECKED"
            If verifs_postponed_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verifs_postponed_checkbox" & ": " & "CHECKED" & vbCr & vbCr
            If verifs_postponed_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verifs_postponed_checkbox" & ": " & "UNCHECKED" & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "adult_cash_count" & ": " & adult_cash_count
            script_run_lowdown = script_run_lowdown & vbCr & "child_cash_count" & ": " & child_cash_count
            If pregnant_caregiver_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "pregnant_caregiver_checkbox" & ": " & "CHECKED"
            If pregnant_caregiver_checkbox = unchecked then script_run_lowdown = script_run_lowdown & vbCr & "pregnant_caregiver_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "adult_snap_count" & ": " & adult_snap_count
            script_run_lowdown = script_run_lowdown & vbCr & "child_snap_count" & ": " & child_snap_count
            script_run_lowdown = script_run_lowdown & vbCr & "adult_emer_count" & ": " & adult_emer_count
            script_run_lowdown = script_run_lowdown & vbCr & "child_emer_count" & ": " & child_emer_count
            script_run_lowdown = script_run_lowdown & vbCr & "EATS" & ": " & EATS
            script_run_lowdown = script_run_lowdown & vbCr & "relationship_detail" & ": " & relationship_detail & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "determined_income" & ": " & determined_income
            script_run_lowdown = script_run_lowdown & vbCr & "income_review_completed" & ": " & income_review_completed
            script_run_lowdown = script_run_lowdown & vbCr & "jobs_income_yn" & ": " & jobs_income_yn
            script_run_lowdown = script_run_lowdown & vbCr & "busi_income_yn" & ": " & busi_income_yn
            script_run_lowdown = script_run_lowdown & vbCr & "unea_income_yn" & ": " & unea_income_yn & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "determined_assets" & ": " & determined_assets
            script_run_lowdown = script_run_lowdown & vbCr & "assets_review_completed" & ": " & assets_review_completed
            script_run_lowdown = script_run_lowdown & vbCr & "cash_amount_yn" & ": " & cash_amount_yn
            script_run_lowdown = script_run_lowdown & vbCr & "bank_account_yn" & ": " & bank_account_yn
            script_run_lowdown = script_run_lowdown & vbCr & "cash_amount" & ": " & cash_amount & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "determined_shel" & ": " & determined_shel
            script_run_lowdown = script_run_lowdown & vbCr & "shel_review_completed" & ": " & shel_review_completed
            script_run_lowdown = script_run_lowdown & vbCr & "rent_amount" & ": " & rent_amount
            script_run_lowdown = script_run_lowdown & vbCr & "lot_rent_amount" & ": " & lot_rent_amount
            script_run_lowdown = script_run_lowdown & vbCr & "mortgage_amount" & ": " & mortgage_amount
            script_run_lowdown = script_run_lowdown & vbCr & "insurance_amount" & ": " & insurance_amount
            script_run_lowdown = script_run_lowdown & vbCr & "tax_amount" & ": " & tax_amount
            script_run_lowdown = script_run_lowdown & vbCr & "room_amount" & ": " & room_amount
            script_run_lowdown = script_run_lowdown & vbCr & "garage_amount" & ": " & garage_amount
            script_run_lowdown = script_run_lowdown & vbCr & "subsidy_amount" & ": " & subsidy_amount & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "determined_utilities" & ": " & determined_utilities
            script_run_lowdown = script_run_lowdown & vbCr & "heat_expense" & ": " & heat_expense
            script_run_lowdown = script_run_lowdown & vbCr & "ac_expense" & ": " & ac_expense
            script_run_lowdown = script_run_lowdown & vbCr & "electric_expense" & ": " & electric_expense
            script_run_lowdown = script_run_lowdown & vbCr & "phone_expense" & ": " & phone_expense
            script_run_lowdown = script_run_lowdown & vbCr & "none_expense" & ": " & none_expense
            script_run_lowdown = script_run_lowdown & vbCr & "all_utilities" & ": " & all_utilities & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "do_we_have_applicant_id" & ": " & do_we_have_applicant_id & vbCr & vbCr

            If CASH_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH_checkbox" & ": " & "CHECKED"
            If CASH_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH_checkbox" & ": " & "UNCHECKED"
            If GRH_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "GRH_checkbox" & ": " & "CHECKED"
            If GRH_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "GRH_checkbox" & ": " & "UNCHECKED"
            If SNAP_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP_checkbox" & ": " & "CHECKED"
            If SNAP_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP_checkbox" & ": " & "UNCHECKED"
            If HC_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "HC_checkbox" & ": " & "CHECKED"
            If HC_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "HC_checkbox" & ": " & "UNCHECKED"
            If EMER_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER_checkbox" & ": " & "CHECKED"
            If EMER_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "multiple_CAF_dates" & ": " & multiple_CAF_dates
            script_run_lowdown = script_run_lowdown & vbCr & "multiple_interview_dates" & ": " & multiple_interview_dates
            script_run_lowdown = script_run_lowdown & vbCr & "adult_cash" & ": " & adult_cash
            script_run_lowdown = script_run_lowdown & vbCr & "family_cash" & ": " & family_cash
            script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_cash" & ": " & the_process_for_cash
            script_run_lowdown = script_run_lowdown & vbCr & "type_of_cash" & ": " & type_of_cash
            script_run_lowdown = script_run_lowdown & vbCr & "cash_recert_mo" & ": " & cash_recert_mo
            script_run_lowdown = script_run_lowdown & vbCr & "cash_recert_yr" & ": " & cash_recert_yr
            script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_grh" & ": " & the_process_for_grh
            script_run_lowdown = script_run_lowdown & vbCr & "grh_recert_mo" & ": " & grh_recert_mo
            script_run_lowdown = script_run_lowdown & vbCr & "grh_recert_yr" & ": " & grh_recert_yr
			script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_snap" & ": " & the_process_for_snap
            script_run_lowdown = script_run_lowdown & vbCr & "snap_recert_mo" & ": " & snap_recert_mo
            script_run_lowdown = script_run_lowdown & vbCr & "snap_recert_yr" & ": " & snap_recert_yr
            script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_hc" & ": " & the_process_for_hc
            script_run_lowdown = script_run_lowdown & vbCr & "hc_recert_mo" & ": " & hc_recert_mo
            script_run_lowdown = script_run_lowdown & vbCr & "hc_recert_yr" & ": " & hc_recert_yr
            script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_emer" & ": " & the_process_for_emer
            script_run_lowdown = script_run_lowdown & vbCr & "type_of_emer" & ": " & type_of_emer
            ' script_run_lowdown = script_run_lowdown & vbCr & "CAF_type" & ": " & CAF_type
			script_run_lowdown = script_run_lowdown & vbCr & "application_processing" & ": " & application_processing
			script_run_lowdown = script_run_lowdown & vbCr & "recert_processing" & ": " & recert_processing
            script_run_lowdown = script_run_lowdown & vbCr & "CAF_datestamp" & ": " & CAF_datestamp
            script_run_lowdown = script_run_lowdown & vbCr & "PROG_CAF_datestamp" & ": " & PROG_CAF_datestamp
            script_run_lowdown = script_run_lowdown & vbCr & "REVW_CAF_datestamp" & ": " & REVW_CAF_datestamp
            script_run_lowdown = script_run_lowdown & vbCr & "interview_date" & ": " & interview_date
            script_run_lowdown = script_run_lowdown & vbCr & "PROG_interview_date" & ": " & PROG_interview_date
            script_run_lowdown = script_run_lowdown & vbCr & "REVW_interview_date" & ": " & REVW_interview_date
            script_run_lowdown = script_run_lowdown & vbCr & "case_details_and_notes_about_process" & ": " & case_details_and_notes_about_process
            script_run_lowdown = script_run_lowdown & vbCr & "SNAP_recert_is_likely_24_months" & ": " & SNAP_recert_is_likely_24_months
            script_run_lowdown = script_run_lowdown & vbCr & "exp_screening_note_found" & ": " & exp_screening_note_found
            script_run_lowdown = script_run_lowdown & vbCr & "interview_required" & ": " & interview_required
            script_run_lowdown = script_run_lowdown & vbCr & "xfs_screening" & ": " & xfs_screening
            script_run_lowdown = script_run_lowdown & vbCr & "xfs_screening_display" & ": " & xfs_screening_display
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_income" & ": " & caf_one_income
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_assets" & ": " & caf_one_assets
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_resources" & ": " & caf_one_resources
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_rent" & ": " & caf_one_rent
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_utilities" & ": " & caf_one_utilities
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_expenses" & ": " & caf_one_expenses
            script_run_lowdown = script_run_lowdown & vbCr & "exp_det_case_note_found" & ": " & exp_det_case_note_found
            script_run_lowdown = script_run_lowdown & vbCr & "snap_exp_yn" & ": " & snap_exp_yn
            script_run_lowdown = script_run_lowdown & vbCr & "snap_denial_date" & ": " & snap_denial_date
            script_run_lowdown = script_run_lowdown & vbCr & "interview_completed_case_note_found" & ": " & interview_completed_case_note_found
            script_run_lowdown = script_run_lowdown & vbCr & "interview_with" & ": " & interview_with
            script_run_lowdown = script_run_lowdown & vbCr & "interview_type" & ": " & interview_type
            script_run_lowdown = script_run_lowdown & vbCr & "verifications_requested_case_note_found" & ": " & verifications_requested_case_note_found
            script_run_lowdown = script_run_lowdown & vbCr & "verifs_needed" & ": " & verifs_needed
            If verif_snap_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_snap_checkbox" & ":" & "CHECKED"
            If verif_cash_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_cash_checkbox" & ":" & "CHECKED"
            If verif_mfip_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_mfip_checkbox" & ":" & "CHECKED"
            If verif_dwp_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_dwp_checkbox" & ":" & "CHECKED"
            If verif_msa_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_msa_checkbox" & ":" & "CHECKED"
            If verif_ga_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_ga_checkbox" & ":" & "CHECKED"
            If verif_grh_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_grh_checkbox" & ":" & "CHECKED"
            If verif_emer_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_emer_checkbox" & ":" & "CHECKED"
            If verif_hc_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_hc_checkbox" & ":" & "CHECKED"
            If verif_snap_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_snap_checkbox" & ":" & "UNCHECKED"
            If verif_cash_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_cash_checkbox" & ":" & "UNCHECKED"
            If verif_mfip_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_mfip_checkbox" & ":" & "UNCHECKED"
            If verif_dwp_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_dwp_checkbox" & ":" & "UNCHECKED"
            If verif_msa_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_msa_checkbox" & ":" & "UNCHECKED"
            If verif_ga_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_ga_checkbox" & ":" & "UNCHECKED"
            If verif_grh_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_grh_checkbox" & ":" & "UNCHECKED"
            If verif_emer_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_emer_checkbox" & ":" & "UNCHECKED"
            If verif_hc_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "verif_hc_checkbox" & ":" & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "caf_qualifying_questions_case_note_found" & ": " & caf_qualifying_questions_case_note_found
            script_run_lowdown = script_run_lowdown & vbCr & "qual_question_one" & ": " & qual_question_one
            script_run_lowdown = script_run_lowdown & vbCr & "qual_memb_one" & ": " & qual_memb_one
            script_run_lowdown = script_run_lowdown & vbCr & "qual_question_two" & ": " & qual_question_two
            script_run_lowdown = script_run_lowdown & vbCr & "qual_memb_two" & ": " & qual_memb_two
            script_run_lowdown = script_run_lowdown & vbCr & "qual_question_three" & ": " & qual_question_three
            script_run_lowdown = script_run_lowdown & vbCr & "qual_memb_three" & ": " & qual_memb_three
            script_run_lowdown = script_run_lowdown & vbCr & "qual_question_four" & ": " & qual_question_four
            script_run_lowdown = script_run_lowdown & vbCr & "qual_memb_four" & ": " & qual_memb_four
            script_run_lowdown = script_run_lowdown & vbCr & "qual_question_five" & ": " & qual_question_five
            script_run_lowdown = script_run_lowdown & vbCr & "qual_memb_five" & ": " & qual_memb_five
            script_run_lowdown = script_run_lowdown & vbCr & "appt_notc_sent_on" & ": " & appt_notc_sent_on
            script_run_lowdown = script_run_lowdown & vbCr & "appt_date_in_note" & ": " & appt_date_in_note & vbCr & vbCr
            For each HH_MEMB in HH_member_array
                script_run_lowdown = script_run_lowdown & vbCr & "HH_member_array" & ": " & HH_MEMB
            Next
            script_run_lowdown = script_run_lowdown & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "addr_line_one" & ": " & addr_line_one
            script_run_lowdown = script_run_lowdown & vbCr & "addr_line_two" & ": " & addr_line_two
            script_run_lowdown = script_run_lowdown & vbCr & "city" & ": " & city
            script_run_lowdown = script_run_lowdown & vbCr & "state" & ": " & state
            script_run_lowdown = script_run_lowdown & vbCr & "zip" & ": " & zip
            script_run_lowdown = script_run_lowdown & vbCr & "addr_county" & ": " & addr_county
            script_run_lowdown = script_run_lowdown & vbCr & "homeless_yn" & ": " & homeless_yn
            script_run_lowdown = script_run_lowdown & vbCr & "reservation_yn" & ": " & reservation_yn
            script_run_lowdown = script_run_lowdown & vbCr & "addr_verif" & ": " & addr_verif
            script_run_lowdown = script_run_lowdown & vbCr & "living_situation" & ": " & living_situation
            script_run_lowdown = script_run_lowdown & vbCr & "addr_eff_date" & ": " & addr_eff_date
            script_run_lowdown = script_run_lowdown & vbCr & "addr_future_date" & ": " & addr_future_date
            script_run_lowdown = script_run_lowdown & vbCr & "mail_line_one" & ": " & mail_line_one
            script_run_lowdown = script_run_lowdown & vbCr & "mail_line_two" & ": " & mail_line_two
            script_run_lowdown = script_run_lowdown & vbCr & "mail_city_line" & ": " & mail_city_line
            script_run_lowdown = script_run_lowdown & vbCr & "mail_state_line" & ": " & mail_state_line
            script_run_lowdown = script_run_lowdown & vbCr & "mail_zip_line" & ": " & mail_zip_line
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_address" & ": " & notes_on_address & vbCr & vbCr
            For the_members = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                box_one_info = ""
                box_two_info = ""
                box_three_info = ""
                box_four_info = ""
                box_five_info = ""
                box_six_info = ""
                box_seven_info = ""
                box_eight_info = ""
                box_nine_info = ""
                If ALL_MEMBERS_ARRAY(include_cash_checkbox, the_members) = checked Then box_one_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(include_snap_checkbox, the_members) = checked Then box_two_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(include_emer_checkbox, the_members) = checked Then box_three_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(count_cash_checkbox, the_members) = checked Then box_four_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(count_snap_checkbox, the_members) = checked Then box_five_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(count_emer_checkbox, the_members) = checked Then box_six_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(pwe_checkbox, the_members) = checked Then box_seven_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(shel_verif_checkbox, the_members) = checked Then box_eight_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(id_required, the_members) = checked Then box_nine_info = "CHECKED"
                If ALL_MEMBERS_ARRAY(include_cash_checkbox, the_members) = unchecked Then box_one_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(include_snap_checkbox, the_members) = unchecked Then box_two_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(include_emer_checkbox, the_members) = unchecked Then box_three_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(count_cash_checkbox, the_members) = unchecked Then box_four_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(count_snap_checkbox, the_members) = unchecked Then box_five_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(count_emer_checkbox, the_members) = unchecked Then box_six_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(pwe_checkbox, the_members) = unchecked Then box_seven_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(shel_verif_checkbox, the_members) = unchecked Then box_eight_info = "UNCHECKED"
                If ALL_MEMBERS_ARRAY(id_required, the_members) = unchecked Then box_nine_info = "UNCHECKED"

                script_run_lowdown = script_run_lowdown & vbCr & "ALL_MEMBERS_ARRAY" & ": " &ALL_MEMBERS_ARRAY(memb_numb, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_name, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_age, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(full_clt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_id_verif, the_members)&"~**^"&box_one_info&"~**^"&box_two_info&"~**^"&box_three_info&"~**^"&box_four_info&"~**^"&box_five_info&"~**^"&box_six_info&"~**^"&_
                ALL_MEMBERS_ARRAY(clt_wreg_status, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_abawd_status, the_members)&"~**^"&box_seven_info&"~**^"&ALL_MEMBERS_ARRAY(numb_abawd_used, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(list_abawd_mo, the_members)&"~**^"&ALL_MEMBERS_ARRAY(first_second_set, the_members)&"~**^"&ALL_MEMBERS_ARRAY(list_second_set, the_members)&"~**^"&ALL_MEMBERS_ARRAY(explain_no_second, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(numb_banked_mo, the_members)&"~**^"&ALL_MEMBERS_ARRAY(clt_abawd_notes, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_exists, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_subsudized, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_shared, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_rent_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_rent_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_lot_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_lot_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_ins_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_ins_verif,the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_tax_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_tax_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_room_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_room_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_room_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_room_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_retro_garage_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_garage_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt,the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, the_members)&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, the_members)&"~**^"&ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, the_members)&"~**^"&ALL_MEMBERS_ARRAY(wreg_exists, the_members)&"~**^"&box_eight_info&"~**^"&_
                ALL_MEMBERS_ARRAY(shel_verif_added, the_members)&"~**^"&ALL_MEMBERS_ARRAY(gather_detail, the_members)&"~**^"&ALL_MEMBERS_ARRAY(id_detail, the_members)&"~**^"&box_nine_info&"~**^"&_
                ALL_MEMBERS_ARRAY(clt_notes, the_members)
            Next
            script_run_lowdown = script_run_lowdown & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "total_shelter_amount" & ": " & total_shelter_amount
            script_run_lowdown = script_run_lowdown & vbCr & "full_shelter_details" & ": " & full_shelter_details
            script_run_lowdown = script_run_lowdown & vbCr & "shelter_details" & ": " & shelter_details
            script_run_lowdown = script_run_lowdown & vbCr & "shelter_details_two" & ": " & shelter_details_two
            script_run_lowdown = script_run_lowdown & vbCr & "shelter_details_three" & ": " & shelter_details_three
            script_run_lowdown = script_run_lowdown & vbCr & "prosp_heat_air" & ": " & prosp_heat_air
            script_run_lowdown = script_run_lowdown & vbCr & "prosp_electric" & ": " & prosp_electric
            script_run_lowdown = script_run_lowdown & vbCr & "prosp_phone" & ": " & prosp_phone
            script_run_lowdown = script_run_lowdown & vbCr & "hest_information" & ": " & hest_information
            script_run_lowdown = script_run_lowdown & vbCr & "ABPS" & ": " & ABPS
            script_run_lowdown = script_run_lowdown & vbCr & "ACCI" & ": " & ACCI
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_acct" & ": " & notes_on_acct
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_acut" & ": " & notes_on_acut
            script_run_lowdown = script_run_lowdown & vbCr & "AREP" & ": " & AREP
            script_run_lowdown = script_run_lowdown & vbCr & "BILS" & ": " & BILS
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_cash" & ": " & notes_on_cash
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_cars" & ": " & notes_on_cars
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_coex" & ": " & notes_on_coex
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_dcex" & ": " & notes_on_dcex
            script_run_lowdown = script_run_lowdown & vbCr & "DIET" & ": " & DIET
            script_run_lowdown = script_run_lowdown & vbCr & "DISA" & ": " & DISA
            script_run_lowdown = script_run_lowdown & vbCr & "EMPS" & ": " & EMPS
            script_run_lowdown = script_run_lowdown & vbCr & "FACI" & ": " & FACI
            script_run_lowdown = script_run_lowdown & vbCr & "FMED" & ": " & FMED
            script_run_lowdown = script_run_lowdown & vbCr & "IMIG" & ": " & IMIG
            script_run_lowdown = script_run_lowdown & vbCr & "INSA" & ": " & INSA & vbCr & vbCr
            For the_jobs = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                box_one_info = ""
                box_two_info = ""
                If ALL_JOBS_PANELS_ARRAY(estimate_only, the_jobs) = checked Then box_one_info = "CHECKED"
                If ALL_JOBS_PANELS_ARRAY(verif_checkbox, the_jobs) = checked Then box_two_info = "CHECKED"
                If ALL_JOBS_PANELS_ARRAY(estimate_only, the_jobs) = unchecked Then box_one_info = "UNCHECKED"
                If ALL_JOBS_PANELS_ARRAY(verif_checkbox, the_jobs) = unchecked Then box_two_info = "UNCHECKED"
                script_run_lowdown = script_run_lowdown & vbCr & "ALL_JOBS_PANELS_ARRAY" & ": " &ALL_JOBS_PANELS_ARRAY(memb_numb, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(panel_instance, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(employer_name, the_jobs)&"~**^"&box_one_info&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_explain, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_code, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(info_month, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(hrly_wage, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(main_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(job_retro_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(job_prosp_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(retro_hours, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(prosp_hours, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_prosp_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_calc_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(EI_case_note, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_calc_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_prosp_income, the_jobs)&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&ALL_JOBS_PANELS_ARRAY(start_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(end_date, the_jobs)&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&box_two_info&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_added, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(budget_explain, the_jobs)
            Next
            script_run_lowdown = script_run_lowdown & vbCr & vbCr
            For the_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                box_one_info = ""
                box_two_info = ""
                box_three_info = ""
                If ALL_BUSI_PANELS_ARRAY(estimate_only, the_busi) = checked Then box_one_info = "CHECKED"
                If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, the_busi) = checked Then box_two_info = "CHECKED"
                If ALL_BUSI_PANELS_ARRAY(verif_checkbox, the_busi) = checked Then box_three_info = "CHECKED"
                If ALL_BUSI_PANELS_ARRAY(estimate_only, the_busi) = unchecked Then box_one_info = "UNCHECKED"
                If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, the_busi) = unchecked Then box_two_info = "UNCHECKED"
                If ALL_BUSI_PANELS_ARRAY(verif_checkbox, the_busi) = unchecked Then box_three_info = "UNCHECKED"
                script_run_lowdown = script_run_lowdown & vbCr & "ALL_BUSI_PANELS_ARRAY" & ": " &ALL_BUSI_PANELS_ARRAY(memb_numb, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(panel_instance, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_type, the_busi)&"~**^"&box_one_info&"~**^"&ALL_BUSI_PANELS_ARRAY(verif_explain, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(calc_method, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(info_month, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(mthd_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_ret_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_pro_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(cash_income_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_ret_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_pro_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(cash_expense_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_ret_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_pro_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(snap_income_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_ret_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_pro_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(snap_expense_verif, the_busi)&"~**^"&box_two_info&"~**^"&ALL_BUSI_PANELS_ARRAY(start_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(end_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_desc, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_structure, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(share_num, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(share_denom, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(partners_in_HH, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(exp_not_allwd, the_busi)&"~**^"&box_three_info&"~**^"&ALL_BUSI_PANELS_ARRAY(verif_added, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(budget_explain, the_busi)
            Next
            script_run_lowdown = script_run_lowdown & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "cit_id" & ": " & cit_id
            script_run_lowdown = script_run_lowdown & vbCr & "other_assets" & ": " & other_assets
            script_run_lowdown = script_run_lowdown & vbCr & "case_changes" & ": " & case_changes
            script_run_lowdown = script_run_lowdown & vbCr & "PREG" & ": " & PREG
            script_run_lowdown = script_run_lowdown & vbCr & "earned_income" & ": " & earned_income
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_rest" & ": " & notes_on_rest
            script_run_lowdown = script_run_lowdown & vbCr & "SCHL" & ": " & SCHL
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_jobs" & ": " & notes_on_jobs
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_cses" & ": " & notes_on_cses
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_time" & ": " & notes_on_time
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_sanction" & ": " & notes_on_sanction
            For the_unea = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                script_run_lowdown = script_run_lowdown & vbCr & "UNEA_INCOME_ARRAY" & ": " &UNEA_INCOME_ARRAY(memb_numb, the_unea)&"~**^"&UNEA_INCOME_ARRAY(panel_instance, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_type, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_month, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_verif, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_prosp_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_retro_amt, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_SNAP_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_pay_freq, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_pic_date_calc, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_start_date, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_notes, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_account_balance, the_unea)&"~**^"&UNEA_INCOME_ARRAY(direct_CS_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_amt, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(disb_CS_arrears_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(direct_CS_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_notes, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(disb_CS_months, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_prosp_budg, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_months, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_budg, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UNEA_RSDI_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_RSDI_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SSI_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SSI_notes, the_unea)&"~**^"&_
                UNEA_INCOME_ARRAY(UC_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(CS_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(SSA_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(calc_button, the_unea)&"~**^"&UNEA_INCOME_ARRAY(budget_notes, the_unea)
            Next
            script_run_lowdown = script_run_lowdown & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_wreg" & ": " & notes_on_wreg
            script_run_lowdown = script_run_lowdown & vbCr & "full_abawd_info" & ": " & full_abawd_info
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_abawd" & ": " & notes_on_abawd
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_abawd_two" & ": " & notes_on_abawd_two
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_abawd_three" & ": " & notes_on_abawd_three
            script_run_lowdown = script_run_lowdown & vbCr & "programs_applied_for" & ": " & programs_applied_for
            If TIKL_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "TIKL_checkbox" & ": " & "CHECKED"
            If TIKL_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "TIKL_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "interview_memb_list" & ": " & interview_memb_list
            script_run_lowdown = script_run_lowdown & vbCr & "shel_memb_list" & ": " & shel_memb_list
            script_run_lowdown = script_run_lowdown & vbCr & "verification_memb_list" & ": " & verification_memb_list
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_busi" & ": " & notes_on_busi & vbCr & vbCr
            'DLG 1
            If Used_Interpreter_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "Used_Interpreter_checkbox" & ": " & "CHECKED"
            If Used_Interpreter_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "Used_Interpreter_checkbox" & ": " & "UNCHECKED"
            ' objTextStream.WriteLine "how_app_rcvd" & "^~^~^~^~^~^~^" & how_app_rcvd
            script_run_lowdown = script_run_lowdown & vbCr & "arep_id_info" & ": " & arep_id_info
            script_run_lowdown = script_run_lowdown & vbCr & "CS_forms_sent_date" & ": " & CS_forms_sent_date
            script_run_lowdown = script_run_lowdown & vbCr & "case_changes" & ": " & case_changes & vbCr & vbCr
            'DLG 5'
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_ssa_income" & ": " & notes_on_ssa_income
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_VA_income" & ": " & notes_on_VA_income
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_WC_income" & ": " & notes_on_WC_income
            script_run_lowdown = script_run_lowdown & vbCr & "other_uc_income_notes" & ": " & other_uc_income_notes
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_other_UNEA" & ": " & notes_on_other_UNEA & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "hest_information" & ": " & hest_information
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_acut" & ": " & notes_on_acut
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_coex" & ": " & notes_on_coex
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_dcex" & ": " & notes_on_dcex
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_other_deduction" & ": " & notes_on_other_deduction
            script_run_lowdown = script_run_lowdown & vbCr & "expense_notes" & ": " & expense_notes
            If address_confirmation_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "address_confirmation_checkbox" & ": " & "CHECKED"
            If address_confirmation_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "address_confirmation_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "manual_total_shelter" & ": " & manual_total_shelter
            script_run_lowdown = script_run_lowdown & vbCr & "manual_amount_used" & ": " & manual_amount_used
            script_run_lowdown = script_run_lowdown & vbCr & "app_month_assets" & ": " & app_month_assets
            If confirm_no_account_panel_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "confirm_no_account_panel_checkbox" & ": " & "CHECKED"
            If confirm_no_account_panel_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "confirm_no_account_panel_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "notes_on_other_assets" & ": " & notes_on_other_assets
            script_run_lowdown = script_run_lowdown & vbCr & "MEDI" & ": " & MEDI
            script_run_lowdown = script_run_lowdown & vbCr & "DISQ" & ": " & DISQ
            'EXP DET'
            script_run_lowdown = script_run_lowdown & vbCr & "full_determination_done" & ": " & full_determination_done & vbCr & vbCr
            ' Call run_expedited_determination_script_functionality(
            script_run_lowdown = script_run_lowdown & vbCr & "xfs_screening" & ": " & xfs_screening
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_income" & ": " & caf_one_income
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_assets" & ": " & caf_one_assets
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_rent" & ": " & caf_one_rent
            script_run_lowdown = script_run_lowdown & vbCr & "caf_one_utilities" & ": " & caf_one_utilities
            script_run_lowdown = script_run_lowdown & vbCr & "determined_income" & ": " & determined_income
            script_run_lowdown = script_run_lowdown & vbCr & "determined_assets" & ": " & determined_assets
            script_run_lowdown = script_run_lowdown & vbCr & "determined_shel" & ": " & determined_shel
            script_run_lowdown = script_run_lowdown & vbCr & "determined_utilities" & ": " & determined_utilities
            script_run_lowdown = script_run_lowdown & vbCr & "calculated_resources" & ": " & calculated_resources
            script_run_lowdown = script_run_lowdown & vbCr & "calculated_expenses" & ": " & calculated_expenses
            script_run_lowdown = script_run_lowdown & vbCr & "calculated_low_income_asset_test" & ": " & calculated_low_income_asset_test
            script_run_lowdown = script_run_lowdown & vbCr & "calculated_resources_less_than_expenses_test" & ": " & calculated_resources_less_than_expenses_test
            script_run_lowdown = script_run_lowdown & vbCr & "is_elig_XFS" & ": " & is_elig_XFS
            script_run_lowdown = script_run_lowdown & vbCr & "approval_date" & ": " & approval_date
            script_run_lowdown = script_run_lowdown & vbCr & "applicant_id_on_file_yn" & ": " & applicant_id_on_file_yn
            script_run_lowdown = script_run_lowdown & vbCr & "applicant_id_through_SOLQ" & ": " & applicant_id_through_SOLQ
            script_run_lowdown = script_run_lowdown & vbCr & "delay_explanation" & ": " & delay_explanation
            script_run_lowdown = script_run_lowdown & vbCr & "snap_denial_date" & ": " & snap_denial_date
            script_run_lowdown = script_run_lowdown & vbCr & "snap_denial_explain" & ": " & snap_denial_explain
            script_run_lowdown = script_run_lowdown & vbCr & "case_assesment_text" & ": " & case_assesment_text
            script_run_lowdown = script_run_lowdown & vbCr & "next_steps_one" & ": " & next_steps_one
            script_run_lowdown = script_run_lowdown & vbCr & "next_steps_two" & ": " & next_steps_two
            script_run_lowdown = script_run_lowdown & vbCr & "next_steps_three" & ": " & next_steps_three
            script_run_lowdown = script_run_lowdown & vbCr & "next_steps_four" & ": " & next_steps_four
            script_run_lowdown = script_run_lowdown & vbCr & "postponed_verifs_yn" & ": " & postponed_verifs_yn
            script_run_lowdown = script_run_lowdown & vbCr & "list_postponed_verifs" & ": " & list_postponed_verifs
            script_run_lowdown = script_run_lowdown & vbCr & "day_30_from_application" & ": " & day_30_from_application
            script_run_lowdown = script_run_lowdown & vbCr & "other_snap_state" & ": " & other_snap_state
            script_run_lowdown = script_run_lowdown & vbCr & "other_state_reported_benefit_end_date" & ": " & other_state_reported_benefit_end_date
            script_run_lowdown = script_run_lowdown & vbCr & "other_state_benefits_openended" & ": " & other_state_benefits_openended
            script_run_lowdown = script_run_lowdown & vbCr & "other_state_contact_yn" & ": " & other_state_contact_yn
            script_run_lowdown = script_run_lowdown & vbCr & "other_state_verified_benefit_end_date" & ": " & other_state_verified_benefit_end_date
            script_run_lowdown = script_run_lowdown & vbCr & "mn_elig_begin_date" & ": " & mn_elig_begin_date
            script_run_lowdown = script_run_lowdown & vbCr & "action_due_to_out_of_state_benefits" & ": " & action_due_to_out_of_state_benefits
            script_run_lowdown = script_run_lowdown & vbCr & "case_has_previously_postponed_verifs_that_prevent_exp_snap" & ": " & case_has_previously_postponed_verifs_that_prevent_exp_snap
            script_run_lowdown = script_run_lowdown & vbCr & "prev_post_verif_assessment_done" & ": " & prev_post_verif_assessment_done
            script_run_lowdown = script_run_lowdown & vbCr & "previous_date_of_application" & ": " & previous_date_of_application
            script_run_lowdown = script_run_lowdown & vbCr & "previous_expedited_package" & ": " & previous_expedited_package
            script_run_lowdown = script_run_lowdown & vbCr & "prev_verifs_mandatory_yn" & ": " & prev_verifs_mandatory_yn
            script_run_lowdown = script_run_lowdown & vbCr & "prev_verif_list" & ": " & prev_verif_list
            script_run_lowdown = script_run_lowdown & vbCr & "curr_verifs_postponed_yn" & ": " & curr_verifs_postponed_yn
            script_run_lowdown = script_run_lowdown & vbCr & "ongoing_snap_approved_yn" & ": " & ongoing_snap_approved_yn
            script_run_lowdown = script_run_lowdown & vbCr & "prev_post_verifs_recvd_yn" & ": " & prev_post_verifs_recvd_yn
            script_run_lowdown = script_run_lowdown & vbCr & "delay_action_due_to_faci" & ": " & delay_action_due_to_faci
            script_run_lowdown = script_run_lowdown & vbCr & "deny_snap_due_to_faci" & ": " & deny_snap_due_to_faci
            script_run_lowdown = script_run_lowdown & vbCr & "faci_review_completed" & ": " & faci_review_completed
            script_run_lowdown = script_run_lowdown & vbCr & "facility_name" & ": " & facility_name
            script_run_lowdown = script_run_lowdown & vbCr & "snap_inelig_faci_yn" & ": " & snap_inelig_faci_yn
            script_run_lowdown = script_run_lowdown & vbCr & "faci_entry_date" & ": " & faci_entry_date
            script_run_lowdown = script_run_lowdown & vbCr & "faci_release_date" & ": " & faci_release_date
            If release_date_unknown_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "release_date_unknown_checkbox" & ": " & "CHECKED"
            If release_date_unknown_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "release_date_unknown_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "release_within_30_days_yn" & ": " & release_within_30_days_yn & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "next_er_month" & ": " & next_er_month
            script_run_lowdown = script_run_lowdown & vbCr & "next_er_year" & ": " & next_er_year
            script_run_lowdown = script_run_lowdown & vbCr & "CAF_status" & ": " & CAF_status
            script_run_lowdown = script_run_lowdown & vbCr & "actions_taken" & ": " & actions_taken & vbCr & vbCr
            If application_signed_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "application_signed_checkbox" & ": " & "CHECKED"
            If application_signed_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "application_signed_checkbox" & ": " & "UNCHECKED"
            If eDRS_sent_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "eDRS_sent_checkbox" & ": " & "CHECKED"
            If eDRS_sent_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "eDRS_sent_checkbox" & ": " & "UNCHECKED"
            If updated_MMIS_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "updated_MMIS_checkbox" & ": " & "CHECKED"
            If updated_MMIS_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "updated_MMIS_checkbox" & ": " & "UNCHECKED"
            If WF1_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "WF1_checkbox" & ": " & "CHECKED"
            If WF1_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "WF1_checkbox" & ": " & "UNCHECKED"
            If Sent_arep_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "Sent_arep_checkbox" & ": " & "CHECKED"
            If Sent_arep_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "Sent_arep_checkbox" & ": " & "UNCHECKED"
            If intake_packet_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "intake_packet_checkbox" & ": " & "CHECKED"
            If intake_packet_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "intake_packet_checkbox" & ": " & "UNCHECKED"
            If IAA_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "IAA_checkbox" & ": " & "CHECKED"
            If IAA_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "IAA_checkbox" & ": " & "UNCHECKED"
            If recert_period_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "recert_period_checkbox" & ": " & "CHECKED"
            If recert_period_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "recert_period_checkbox" & ": " & "UNCHECKED"
            If R_R_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "R_R_checkbox" & ": " & "CHECKED"
            If R_R_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "R_R_checkbox" & ": " & "UNCHECKED"
            If E_and_T_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "E_and_T_checkbox" & ": " & "CHECKED"
            If E_and_T_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "E_and_T_checkbox" & ": " & "UNCHECKED"
            If elig_req_explained_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "elig_req_explained_checkbox" & ": " & "CHECKED"
            If elig_req_explained_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "elig_req_explained_checkbox" & ": " & "UNCHECKED"
            If benefit_payment_explained_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "benefit_payment_explained_checkbox" & ": " & "CHECKED"
            If benefit_payment_explained_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "benefit_payment_explained_checkbox" & ": " & "UNCHECKED"
            script_run_lowdown = script_run_lowdown & vbCr & "other_notes" & ": " & other_notes
            If client_delay_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "client_delay_checkbox" & ": " & "CHECKED"
            If client_delay_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "client_delay_checkbox" & ": " & "UNCHECKED"
            If TIKL_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "TIKL_checkbox" & ": " & "CHECKED"
            If TIKL_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "TIKL_checkbox" & ": " & "UNCHECKED"
            If client_delay_TIKL_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "client_delay_TIKL_checkbox" & ": " & "CHECKED" & vbCr & vbCr
            If client_delay_TIKL_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "client_delay_TIKL_checkbox" & ": " & "UNCHECKED" & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "verif_req_form_sent_date" & ": " & verif_req_form_sent_date
            script_run_lowdown = script_run_lowdown & vbCr & "worker_signature" & ": " & worker_signature
            script_run_lowdown = script_run_lowdown & vbCr & "script_information_was_restored" & ": " & script_information_was_restored
            case_note_lowdown = replace(case_notes_information, "%^%", vbCr)
            script_run_lowdown = script_run_lowdown & vbCr & vbCr & case_note_lowdown
        End If
    End With
end function

function restore_your_work(vars_filled)
'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run

	'Now determines name of file
	local_changelog_path = user_myDocs_folder & "caf-variables-" & MAXIS_case_number & "-info.txt"
    script_information_was_restored = True

	With objFSO

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(local_changelog_path) = True then

			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

			If pull_variables = vbYes Then
                vars_filled = True
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(local_changelog_path, ForReading)

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				saved_caf_details = split(every_line_in_text_file, vbNewLine)

                known_ref = 0
                known_membs = 0
                known_jobs = 0
                known_busi = 0
                known_unea = 0

				CASH_checkbox = unchecked
				GRH_checkbox = unchecked
				SNAP_checkbox = unchecked
				HC_checkbox = unchecked
				EMER_checkbox = unchecked

                For Each text_line in saved_caf_details										'read each line in the file
                    If Instr(text_line, "^~^~^~^~^~^~^") <> 0 Then
                        line_info = split(text_line, "^~^~^~^~^~^~^")								'creating a small array for each line. 0 has the header and 1 has the information
                        line_info(0) = trim(line_info(0))
            			'here we add the information from TXT to Excel
                        If line_info(0) = "" Then variable = line_info(1)

                        If line_info(0) = "MAXIS_footer_month" Then MAXIS_footer_month = line_info(1)
                        If line_info(0) = "MAXIS_footer_year" Then MAXIS_footer_year = line_info(1)
                        If line_info(0) = "CAF_form" Then CAF_form = line_info(1)


                        If line_info(0) = "number_verifs_checkbox" and line_info(1) = "CHECKED" Then number_verifs_checkbox = checked
                        If line_info(0) = "verifs_postponed_checkbox" and line_info(1) = "CHECKED" Then verifs_postponed_checkbox = checked

                        If line_info(0) = "adult_cash_count" Then adult_cash_count = line_info(1)
                        If line_info(0) = "child_cash_count" Then child_cash_count = line_info(1)
                        If line_info(0) = "pregnant_caregiver_checkbox" and line_info(1) = "CHECKED" Then pregnant_caregiver_checkbox = checked
                        If line_info(0) = "adult_snap_count" Then adult_snap_count = line_info(1)
                        If line_info(0) = "child_snap_count" Then child_snap_count = line_info(1)
                        If line_info(0) = "adult_emer_count" Then adult_emer_count = line_info(1)
                        If line_info(0) = "child_emer_count" Then child_emer_count = line_info(1)
                        If line_info(0) = "EATS" Then EATS = line_info(1)
                        If line_info(0) = "relationship_detail" Then relationship_detail = line_info(1)
                        If line_info(0) = "determined_income" Then determined_income = line_info(1)
                        If line_info(0) = "income_review_completed" Then income_review_completed = line_info(1)
                        If UCase(income_review_completed) = "TRUE" Then income_review_completed = True
                        If UCase(income_review_completed) = "FALSE" Then income_review_completed = False
                        If line_info(0) = "jobs_income_yn" Then jobs_income_yn = line_info(1)
                        If line_info(0) = "busi_income_yn" Then busi_income_yn = line_info(1)
                        If line_info(0) = "unea_income_yn" Then unea_income_yn = line_info(1)
                        ' If line_info(0) = "determined_assets" Then determined_assets = line_info(1)
                        If line_info(0) = "assets_review_completed" Then assets_review_completed = line_info(1)
                        If UCase(assets_review_completed) = "TRUE" Then assets_review_completed = True
                        If UCase(assets_review_completed) = "FALSE" Then assets_review_completed = False
                        If line_info(0) = "cash_amount_yn" Then cash_amount_yn = line_info(1)
                        If line_info(0) = "bank_account_yn" Then bank_account_yn = line_info(1)
                        If line_info(0) = "cash_amount" Then cash_amount = line_info(1)
                        ' If line_info(0) = "determined_shel" Then determined_shel = line_info(1)
                        If line_info(0) = "shel_review_completed" Then shel_review_completed = line_info(1)
                        If UCase(shel_review_completed) = "TRUE" Then shel_review_completed = True
                        If UCase(shel_review_completed) = "FALSE" Then shel_review_completed = False
                        If line_info(0) = "rent_amount" Then rent_amount = line_info(1)
                        If line_info(0) = "lot_rent_amount" Then lot_rent_amount = line_info(1)
                        If line_info(0) = "mortgage_amount" Then mortgage_amount = line_info(1)
                        If line_info(0) = "insurance_amount" Then insurance_amount = line_info(1)
                        If line_info(0) = "tax_amount" Then tax_amount = line_info(1)
                        If line_info(0) = "room_amount" Then room_amount = line_info(1)
                        If line_info(0) = "garage_amount" Then garage_amount = line_info(1)
                        If line_info(0) = "subsidy_amount" Then subsidy_amount = line_info(1)
                        ' If line_info(0) = "determined_utilities" Then determined_utilities = line_info(1)
                        If line_info(0) = "heat_expense" Then heat_expense = line_info(1)
                        If UCase(heat_expense) = "TRUE" Then heat_expense = True
                        If UCase(heat_expense) = "FALSE" Then heat_expense = False
                        If line_info(0) = "ac_expense" Then ac_expense = line_info(1)
                        If UCase(ac_expense) = "TRUE" Then ac_expense = True
                        If UCase(ac_expense) = "FALSE" Then ac_expense = False
                        If line_info(0) = "electric_expense" Then electric_expense = line_info(1)
                        If UCase(electric_expense) = "TRUE" Then electric_expense = True
                        If UCase(electric_expense) = "FALSE" Then electric_expense = False
                        If line_info(0) = "phone_expense" Then phone_expense = line_info(1)
                        If UCase(phone_expense) = "TRUE" Then phone_expense = True
                        If UCase(phone_expense) = "FALSE" Then phone_expense = False
                        If line_info(0) = "none_expense" Then none_expense = line_info(1)
                        If UCase(none_expense) = "TRUE" Then none_expense = True
                        If UCase(none_expense) = "FALSE" Then none_expense = False
                        If line_info(0) = "all_utilities" Then all_utilities = line_info(1)
                        If line_info(0) = "do_we_have_applicant_id" Then do_we_have_applicant_id = line_info(1)
                        If UCase(do_we_have_applicant_id) = "TRUE" Then do_we_have_applicant_id = True
                        If UCase(do_we_have_applicant_id) = "FALSE" Then do_we_have_applicant_id = False
                        If line_info(0) = "adult_cash" Then adult_cash = line_info(1)
                        If UCase(adult_cash) = "TRUE" Then adult_cash = True
                        If UCase(adult_cash) = "FALSE" Then adult_cash = False
                        If line_info(0) = "family_cash" Then family_cash = line_info(1)
                        If UCase(family_cash) = "TRUE" Then family_cash = True
                        If UCase(family_cash) = "FALSE" Then family_cash = False
						If line_info(0) = "multiple_CAF_dates" Then multiple_CAF_dates = line_info(1)
                        If UCase(multiple_CAF_dates) = "TRUE" Then multiple_CAF_dates = True
                        If UCase(multiple_CAF_dates) = "FALSE" Then multiple_CAF_dates = False
						If line_info(0) = "multiple_interview_dates" Then multiple_interview_dates = line_info(1)
                        If UCase(multiple_interview_dates) = "TRUE" Then multiple_interview_dates = True
                        If UCase(multiple_interview_dates) = "FALSE" Then multiple_interview_dates = False

                        If line_info(0) = "CASH_checkbox" and line_info(1) = "CHECKED" Then CASH_checkbox = checked
                        If line_info(0) = "GRH_checkbox" and line_info(1) = "CHECKED" Then GRH_checkbox = checked
                        If line_info(0) = "SNAP_checkbox" and line_info(1) = "CHECKED" Then SNAP_checkbox = checked
                        If line_info(0) = "HC_checkbox" and line_info(1) = "CHECKED" Then HC_checkbox = checked
                        If line_info(0) = "EMER_checkbox" and line_info(1) = "CHECKED" Then EMER_checkbox = checked

                        If line_info(0) = "the_process_for_cash" Then the_process_for_cash = line_info(1)
                        If line_info(0) = "type_of_cash" Then type_of_cash = line_info(1)
                        If line_info(0) = "cash_recert_mo" Then cash_recert_mo = line_info(1)
                        If line_info(0) = "cash_recert_yr" Then cash_recert_yr = line_info(1)
                        If line_info(0) = "the_process_for_grh" Then the_process_for_grh = line_info(1)
                        If line_info(0) = "grh_recert_mo" Then grh_recert_mo = line_info(1)
                        If line_info(0) = "grh_recert_yr" Then grh_recert_yr = line_info(1)
						If line_info(0) = "the_process_for_snap" Then the_process_for_snap = line_info(1)
                        If line_info(0) = "snap_recert_mo" Then snap_recert_mo = line_info(1)
                        If line_info(0) = "snap_recert_yr" Then snap_recert_yr = line_info(1)
                        If line_info(0) = "the_process_for_hc" Then the_process_for_hc = line_info(1)
                        If line_info(0) = "hc_recert_mo" Then hc_recert_mo = line_info(1)
                        If line_info(0) = "hc_recert_yr" Then hc_recert_yr = line_info(1)
						If line_info(0) = "the_process_for_emer" Then the_process_for_emer = line_info(1)
                        If line_info(0) = "type_of_emer" Then type_of_emer = line_info(1)
                        ' If line_info(0) = "CAF_type" Then CAF_type = line_info(1)
						If line_info(0) = "application_processing" Then application_processing = line_info(1)
						If UCase(application_processing) = "TRUE" Then application_processing = True
                        If UCase(application_processing) = "FALSE" Then application_processing = False
						If line_info(0) = "recert_processing" Then recert_processing = line_info(1)
                        If UCase(recert_processing) = "TRUE" Then recert_processing = True
                        If UCase(recert_processing) = "FALSE" Then recert_processing = False
						If line_info(0) = "CAF_datestamp" Then CAF_datestamp = line_info(1)
                        If line_info(0) = "PROG_CAF_datestamp" Then PROG_CAF_datestamp = line_info(1)
                        If line_info(0) = "REVW_CAF_datestamp" Then REVW_CAF_datestamp = line_info(1)
                        If line_info(0) = "interview_date" Then interview_date = line_info(1)
                        If line_info(0) = "PROG_interview_date" Then PROG_interview_date = line_info(1)
                        If line_info(0) = "REVW_interview_date" Then REVW_interview_date = line_info(1)
                        If line_info(0) = "case_details_and_notes_about_process" Then case_details_and_notes_about_process = line_info(1)
                        If line_info(0) = "SNAP_recert_is_likely_24_months" Then SNAP_recert_is_likely_24_months = line_info(1)
                        If UCase(SNAP_recert_is_likely_24_months) = "TRUE" Then SNAP_recert_is_likely_24_months = True
                        If UCase(SNAP_recert_is_likely_24_months) = "FALSE" Then SNAP_recert_is_likely_24_months = False
                        If line_info(0) = "exp_screening_note_found" Then exp_screening_note_found = line_info(1)
                        If UCase(exp_screening_note_found) = "TRUE" Then exp_screening_note_found = True
                        If UCase(exp_screening_note_found) = "FALSE" Then exp_screening_note_found = False
                        If line_info(0) = "interview_required" Then interview_required = line_info(1)
                        If UCase(interview_required) = "TRUE" Then interview_required = True
                        If UCase(interview_required) = "FALSE" Then interview_required = False
                        If line_info(0) = "xfs_screening" Then xfs_screening = line_info(1)
                        If line_info(0) = "xfs_screening_display" Then xfs_screening_display = line_info(1)
                        If line_info(0) = "caf_one_income" Then caf_one_income = line_info(1)
                        If line_info(0) = "caf_one_assets" Then caf_one_assets = line_info(1)
                        If line_info(0) = "caf_one_resources" Then caf_one_resources = line_info(1)
                        If line_info(0) = "caf_one_rent" Then caf_one_rent = line_info(1)
                        If line_info(0) = "caf_one_utilities" Then caf_one_utilities = line_info(1)
                        If line_info(0) = "caf_one_expenses" Then caf_one_expenses = line_info(1)
                        If line_info(0) = "exp_det_case_note_found" Then exp_det_case_note_found = line_info(1)
                        If UCase(exp_det_case_note_found) = "TRUE" Then exp_det_case_note_found = True
                        If UCase(exp_det_case_note_found) = "FALSE" Then exp_det_case_note_found = False
                        If line_info(0) = "snap_exp_yn" Then snap_exp_yn = line_info(1)
                        If line_info(0) = "snap_denial_date" Then snap_denial_date = line_info(1)
                        If line_info(0) = "interview_completed_case_note_found" Then interview_completed_case_note_found = line_info(1)
                        If UCase(interview_completed_case_note_found) = "TRUE" Then interview_completed_case_note_found = True
                        If UCase(interview_completed_case_note_found) = "FALSE" Then interview_completed_case_note_found = False
                        If line_info(0) = "interview_with" Then interview_with = line_info(1)
                        If line_info(0) = "interview_type" Then interview_type = line_info(1)
                        If line_info(0) = "verifications_requested_case_note_found" Then verifications_requested_case_note_found = line_info(1)
                        If UCase(verifications_requested_case_note_found) = "TRUE" Then verifications_requested_case_note_found = True
                        If UCase(verifications_requested_case_note_found) = "FALSE" Then verifications_requested_case_note_found = False
                        If line_info(0) = "verifs_needed" Then verifs_needed = line_info(1)
                        If line_info(0) = "verif_snap_checkbox" and line_info(1) = "CHECKED" then verif_snap_checkbox = checked
                        If line_info(0) = "verif_cash_checkbox" and line_info(1) = "CHECKED" then verif_cash_checkbox = checked
                        If line_info(0) = "verif_mfip_checkbox" and line_info(1) = "CHECKED" then verif_mfip_checkbox = checked
                        If line_info(0) = "verif_dwp_checkbox" and line_info(1) = "CHECKED" then verif_dwp_checkbox = checked
                        If line_info(0) = "verif_msa_checkbox" and line_info(1) = "CHECKED" then verif_msa_checkbox = checked
                        If line_info(0) = "verif_ga_checkbox" and line_info(1) = "CHECKED" then verif_ga_checkbox = checked
                        If line_info(0) = "verif_grh_checkbox" and line_info(1) = "CHECKED" then verif_grh_checkbox = checked
                        If line_info(0) = "verif_emer_checkbox" and line_info(1) = "CHECKED" then verif_emer_checkbox = checked
                        If line_info(0) = "verif_hc_checkbox" and line_info(1) = "CHECKED" then verif_hc_checkbox = checked
                        If line_info(0) = "caf_qualifying_questions_case_note_found" Then caf_qualifying_questions_case_note_found = line_info(1)
                        If UCase(caf_qualifying_questions_case_note_found) = "TRUE" Then caf_qualifying_questions_case_note_found = True
                        If UCase(caf_qualifying_questions_case_note_found) = "FALSE" Then caf_qualifying_questions_case_note_found = False
                        If line_info(0) = "qual_question_one" Then qual_question_one = line_info(1)
                        If line_info(0) = "qual_memb_one" Then qual_memb_one = line_info(1)
                        If line_info(0) = "qual_question_two" Then qual_question_two = line_info(1)
                        If line_info(0) = "qual_memb_two" Then qual_memb_two = line_info(1)
                        If line_info(0) = "qual_question_three" Then qual_question_three = line_info(1)
                        If line_info(0) = "qual_memb_three" Then qual_memb_three = line_info(1)
                        If line_info(0) = "qual_question_four" Then qual_question_four = line_info(1)
                        If line_info(0) = "qual_memb_four" Then qual_memb_four = line_info(1)
                        If line_info(0) = "qual_question_five" Then qual_question_five = line_info(1)
                        If line_info(0) = "qual_memb_five" Then qual_memb_five = line_info(1)
                        If line_info(0) = "appt_notc_sent_on" Then appt_notc_sent_on = line_info(1)
                        If line_info(0) = "appt_date_in_note" Then appt_date_in_note = line_info(1)
                        If line_info(0) = "HH_member_array" Then
                            ReDim Preserve HH_member_array(known_ref)
                            HH_member_array(known_ref) = line_info(1)
                            known_ref = known_ref + 1
                        End If
                        If line_info(0) = "addr_line_one" Then addr_line_one = line_info(1)
                        If line_info(0) = "addr_line_two" Then addr_line_two = line_info(1)
                        If line_info(0) = "city" Then city = line_info(1)
                        If line_info(0) = "state" Then state = line_info(1)
                        If line_info(0) = "zip" Then zip = line_info(1)
                        If line_info(0) = "addr_county" Then addr_county = line_info(1)
                        If line_info(0) = "homeless_yn" Then homeless_yn = line_info(1)
                        If line_info(0) = "reservation_yn" Then reservation_yn = line_info(1)
                        If line_info(0) = "addr_verif" Then addr_verif = line_info(1)
                        If line_info(0) = "living_situation" Then living_situation = line_info(1)
                        If line_info(0) = "addr_eff_date" Then addr_eff_date = line_info(1)
                        If line_info(0) = "addr_future_date" Then addr_future_date = line_info(1)
                        If line_info(0) = "mail_line_one" Then mail_line_one = line_info(1)
                        If line_info(0) = "mail_line_two" Then mail_line_two = line_info(1)
                        If line_info(0) = "mail_city_line" Then mail_city_line = line_info(1)
                        If line_info(0) = "mail_state_line" Then mail_state_line = line_info(1)
                        If line_info(0) = "mail_zip_line" Then mail_zip_line = line_info(1)
                        If line_info(0) = "notes_on_address" Then notes_on_address = line_info(1)

                        If line_info(0) = "ALL_MEMBERS_ARRAY" Then
                            array_info = line_info(1)
                            array_info = split(array_info, "~**^")
                            ReDim Preserve ALL_MEMBERS_ARRAY(clt_notes, known_membs)

                            ALL_MEMBERS_ARRAY(memb_numb, known_membs)                   = array_info(0)
                            ALL_MEMBERS_ARRAY(clt_name, known_membs)                    = array_info(1)
                            ALL_MEMBERS_ARRAY(clt_age, known_membs)                     = array_info(2)
                            ALL_MEMBERS_ARRAY(full_clt, known_membs)                    = array_info(3)
                            ALL_MEMBERS_ARRAY(clt_id_verif, known_membs)                = array_info(4)
                            If array_info(5) = "CHECKED" Then ALL_MEMBERS_ARRAY(include_cash_checkbox, known_membs) = checked
                            If array_info(6) = "CHECKED" Then ALL_MEMBERS_ARRAY(include_snap_checkbox, known_membs) = checked
                            If array_info(7) = "CHECKED" Then ALL_MEMBERS_ARRAY(include_emer_checkbox, known_membs) = checked
                            If array_info(8) = "CHECKED" Then ALL_MEMBERS_ARRAY(count_cash_checkbox, known_membs) = checked
                            If array_info(9) = "CHECKED" Then ALL_MEMBERS_ARRAY(count_snap_checkbox, known_membs) = checked
                            If array_info(10) = "CHECKED" Then ALL_MEMBERS_ARRAY(count_emer_checkbox, known_membs) = checked
                            ALL_MEMBERS_ARRAY(clt_wreg_status, known_membs)             = array_info(11)
                            ALL_MEMBERS_ARRAY(clt_abawd_status, known_membs)            = array_info(12)
                            If array_info(13) = "CHECKED" Then ALL_MEMBERS_ARRAY(include_cash_checkbox, known_membs) = checked
                            ALL_MEMBERS_ARRAY(numb_abawd_used, known_membs)             = array_info(14)
                            ALL_MEMBERS_ARRAY(list_abawd_mo, known_membs)               = array_info(15)
                            ALL_MEMBERS_ARRAY(first_second_set, known_membs)            = array_info(16)
                            ALL_MEMBERS_ARRAY(list_second_set, known_membs)             = array_info(17)
                            ALL_MEMBERS_ARRAY(explain_no_second, known_membs)           = array_info(18)
                            ALL_MEMBERS_ARRAY(numb_banked_mo, known_membs)              = array_info(19)
                            ALL_MEMBERS_ARRAY(clt_abawd_notes, known_membs)             = array_info(20)
                            ALL_MEMBERS_ARRAY(shel_exists, known_membs)                 = array_info(21)
                            If UCase(ALL_MEMBERS_ARRAY(shel_exists, known_membs)) = "TRUE" Then ALL_MEMBERS_ARRAY(shel_exists, known_membs) = True
                            If UCase(ALL_MEMBERS_ARRAY(shel_exists, known_membs)) = "FALSE" Then ALL_MEMBERS_ARRAY(shel_exists, known_membs) = False
                            ALL_MEMBERS_ARRAY(shel_subsudized, known_membs)             = array_info(22)
                            ALL_MEMBERS_ARRAY(shel_shared, known_membs)                 = array_info(23)
                            ALL_MEMBERS_ARRAY(shel_retro_rent_amt, known_membs)         = array_info(24)
                            ALL_MEMBERS_ARRAY(shel_retro_rent_verif, known_membs)       = array_info(25)
                            ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, known_membs)         = array_info(26)
                            ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, known_membs)       = array_info(27)
                            ALL_MEMBERS_ARRAY(shel_retro_lot_amt, known_membs)          = array_info(28)
                            ALL_MEMBERS_ARRAY(shel_retro_lot_verif, known_membs)        = array_info(29)
                            ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, known_membs)          = array_info(30)
                            ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, known_membs)        = array_info(31)
                            ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, known_membs)     = array_info(32)
                            ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, known_membs)   = array_info(33)
                            ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, known_membs)     = array_info(34)
                            ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, known_membs)   = array_info(35)
                            ALL_MEMBERS_ARRAY(shel_retro_ins_amt, known_membs)          = array_info(36)
                            ALL_MEMBERS_ARRAY(shel_retro_ins_verif,known_membs)         = array_info(37)
                            ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, known_membs)          = array_info(38)
                            ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, known_membs)        = array_info(39)
                            ALL_MEMBERS_ARRAY(shel_retro_tax_amt, known_membs)          = array_info(40)
                            ALL_MEMBERS_ARRAY(shel_retro_tax_verif, known_membs)        = array_info(41)
                            ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, known_membs)          = array_info(42)
                            ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, known_membs)        = array_info(43)
                            ALL_MEMBERS_ARRAY(shel_retro_room_amt, known_membs)         = array_info(44)
                            ALL_MEMBERS_ARRAY(shel_retro_room_verif, known_membs)       = array_info(45)
                            ALL_MEMBERS_ARRAY(shel_prosp_room_amt, known_membs)         = array_info(46)
                            ALL_MEMBERS_ARRAY(shel_prosp_room_verif, known_membs)       = array_info(47)
                            ALL_MEMBERS_ARRAY(shel_retro_garage_amt, known_membs)       = array_info(48)
                            ALL_MEMBERS_ARRAY(shel_retro_garage_verif, known_membs)     = array_info(49)
                            ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, known_membs)       = array_info(50)
                            ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, known_membs)     = array_info(51)
                            ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt,known_membs)       = array_info(52)
                            ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, known_membs)    = array_info(53)
                            ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, known_membs)      = array_info(54)
                            ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, known_membs)    = array_info(55)
                            ALL_MEMBERS_ARRAY(wreg_exists, known_membs)                 = array_info(56)
                            If UCase(ALL_MEMBERS_ARRAY(wreg_exists, known_membs)) = "TRUE" Then ALL_MEMBERS_ARRAY(wreg_exists, known_membs) = True
                            If UCase(ALL_MEMBERS_ARRAY(wreg_exists, known_membs)) = "FALSE" Then ALL_MEMBERS_ARRAY(wreg_exists, known_membs) = False
                            If array_info(57) = "CHECKED" Then ALL_MEMBERS_ARRAY(include_cash_checkbox, known_membs) = checked
                            ALL_MEMBERS_ARRAY(shel_verif_added, known_membs)            = array_info(58)
                            If UCase(ALL_MEMBERS_ARRAY(shel_verif_added, known_membs)) = "TRUE" Then ALL_MEMBERS_ARRAY(shel_verif_added, known_membs) = True
                            If UCase(ALL_MEMBERS_ARRAY(shel_verif_added, known_membs)) = "FALSE" Then ALL_MEMBERS_ARRAY(shel_verif_added, known_membs) = False
                            ALL_MEMBERS_ARRAY(gather_detail, known_membs)               = array_info(59)
                            If UCase(ALL_MEMBERS_ARRAY(gather_detail, known_membs)) = "TRUE" Then ALL_MEMBERS_ARRAY(gather_detail, known_membs) = True
                            If UCase(ALL_MEMBERS_ARRAY(gather_detail, known_membs)) = "FALSE" Then ALL_MEMBERS_ARRAY(gather_detail, known_membs) = False
                            ALL_MEMBERS_ARRAY(id_detail, known_membs)                   = array_info(60)
                            If array_info(61) = "CHECKED" Then ALL_MEMBERS_ARRAY(id_required, known_membs) = checked
                            ALL_MEMBERS_ARRAY(clt_notes, known_membs)                   = array_info(62)
                            known_membs = known_membs + 1
                        End If
                        If line_info(0) = "total_shelter_amount" Then total_shelter_amount = line_info(1)
                        If line_info(0) = "full_shelter_details" Then full_shelter_details = line_info(1)
                        If line_info(0) = "shelter_details" Then shelter_details = line_info(1)
                        If line_info(0) = "shelter_details_two" Then shelter_details_two = line_info(1)
                        If line_info(0) = "shelter_details_three" Then shelter_details_three = line_info(1)
                        If line_info(0) = "prosp_heat_air" Then prosp_heat_air = line_info(1)
                        If line_info(0) = "prosp_electric" Then prosp_electric = line_info(1)
                        If line_info(0) = "prosp_phone" Then prosp_phone = line_info(1)
                        If line_info(0) = "hest_information" Then hest_information = line_info(1)
                        If line_info(0) = "ABPS" Then ABPS = line_info(1)
                        If line_info(0) = "ACCI" Then ACCI = line_info(1)
                        If line_info(0) = "notes_on_acct" Then notes_on_acct = line_info(1)
                        If line_info(0) = "notes_on_acut" Then notes_on_acut = line_info(1)
                        If line_info(0) = "AREP" Then AREP = line_info(1)
                        If line_info(0) = "BILS" Then BILS = line_info(1)
                        If line_info(0) = "notes_on_cash" Then notes_on_cash = line_info(1)
                        If line_info(0) = "notes_on_cars" Then notes_on_cars = line_info(1)
                        If line_info(0) = "notes_on_coex" Then notes_on_coex = line_info(1)
                        If line_info(0) = "notes_on_dcex" Then notes_on_dcex = line_info(1)
                        If line_info(0) = "DIET" Then DIET = line_info(1)
                        If line_info(0) = "DISA" Then DISA = line_info(1)
                        If line_info(0) = "EMPS" Then EMPS = line_info(1)
                        If line_info(0) = "FACI" Then FACI = line_info(1)
                        If line_info(0) = "FMED" Then FMED = line_info(1)
                        If line_info(0) = "IMIG" Then IMIG = line_info(1)
                        If line_info(0) = "INSA" Then INSA = line_info(1)
                        If line_info(0) = "ALL_JOBS_PANELS_ARRAY" Then
                            array_info = line_info(1)
                            array_info = split(array_info, "~**^")
                            ReDim Preserve ALL_JOBS_PANELS_ARRAY(budget_explain, known_jobs)

                            ALL_JOBS_PANELS_ARRAY(memb_numb, known_jobs)            = array_info(0)
                            ALL_JOBS_PANELS_ARRAY(panel_instance, known_jobs)       = array_info(1)
                            ALL_JOBS_PANELS_ARRAY(employer_name, known_jobs)        = array_info(2)
                            If array_info(3) = "CHECKED" Then ALL_JOBS_PANELS_ARRAY(estimate_only, known_jobs) = checked
                            ALL_JOBS_PANELS_ARRAY(verif_explain, known_jobs)        = array_info(4)
                            ALL_JOBS_PANELS_ARRAY(verif_code, known_jobs)           = array_info(5)
                            ALL_JOBS_PANELS_ARRAY(info_month, known_jobs)           = array_info(6)
                            ALL_JOBS_PANELS_ARRAY(hrly_wage, known_jobs)            = array_info(7)
                            ALL_JOBS_PANELS_ARRAY(main_pay_freq, known_jobs)        = array_info(8)
                            ALL_JOBS_PANELS_ARRAY(job_retro_income, known_jobs)     = array_info(9)
                            ALL_JOBS_PANELS_ARRAY(job_prosp_income, known_jobs)     = array_info(10)
                            ALL_JOBS_PANELS_ARRAY(retro_hours, known_jobs)          = array_info(11)
                            ALL_JOBS_PANELS_ARRAY(prosp_hours, known_jobs)          = array_info(12)
                            ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, known_jobs)  = array_info(13)
                            ALL_JOBS_PANELS_ARRAY(pic_pay_freq, known_jobs)         = array_info(14)
                            ALL_JOBS_PANELS_ARRAY(pic_prosp_income, known_jobs)     = array_info(15)
                            ALL_JOBS_PANELS_ARRAY(pic_calc_date, known_jobs)        = array_info(16)
                            ALL_JOBS_PANELS_ARRAY(EI_case_note, known_jobs)         = array_info(17)
                            ALL_JOBS_PANELS_ARRAY(grh_calc_date, known_jobs)        = array_info(18)
                            ALL_JOBS_PANELS_ARRAY(grh_pay_freq, known_jobs)         = array_info(19)
                            ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, known_jobs)   = array_info(20)
                            ALL_JOBS_PANELS_ARRAY(grh_prosp_income, known_jobs)     = array_info(21)
                            ALL_JOBS_PANELS_ARRAY(start_date, known_jobs)           = array_info(25)
                            ALL_JOBS_PANELS_ARRAY(end_date, known_jobs)             = array_info(26)
                            If array_info(33) = "CHECKED" Then ALL_JOBS_PANELS_ARRAY(verif_checkbox, known_jobs) = checked
                            ALL_JOBS_PANELS_ARRAY(verif_added, known_jobs)          = array_info(34)
                            If UCase(ALL_JOBS_PANELS_ARRAY(verif_added, known_jobs)) = "TRUE" Then ALL_JOBS_PANELS_ARRAY(verif_added, known_jobs) = True
                            If UCase(ALL_JOBS_PANELS_ARRAY(verif_added, known_jobs)) = "FALSE" Then ALL_JOBS_PANELS_ARRAY(verif_added, known_jobs) = False
                            ALL_JOBS_PANELS_ARRAY(budget_explain, known_jobs)       = array_info(35)

                            known_jobs = known_jobs + 1
                        End If
                        ' For the_jobs = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                        '     known_jobs = 0
                        '     box_one_info = ""
                        '     box_two_info = ""
                        '     If ALL_JOBS_PANELS_ARRAY(estimate_only, the_jobs) = checked Then box_one_info = "CHECKED"
                        '     If ALL_JOBS_PANELS_ARRAY(verif_checkbox, the_jobs) = checked Then box_two_info = "CHECKED"
                        '     objTextStream.WriteLine "ALL_JOBS_PANELS_ARRAY" & "^~^~^~^~^~^~^" &ALL_JOBS_PANELS_ARRAY(memb_numb, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(panel_instance, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(employer_name, the_jobs)&"~**^"&box_one_info&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_explain, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_code, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(info_month, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(hrly_wage, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(main_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(job_retro_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(job_prosp_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(retro_hours, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(prosp_hours, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_prosp_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(pic_calc_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(EI_case_note, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_calc_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_pay_freq, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(grh_prosp_income, the_jobs)&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&ALL_JOBS_PANELS_ARRAY(start_date, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(end_date, the_jobs)&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&""&"~**^"&box_two_info&"~**^"&ALL_JOBS_PANELS_ARRAY(verif_added, the_jobs)&"~**^"&ALL_JOBS_PANELS_ARRAY(budget_explain, the_jobs)
                        ' Next
                        If line_info(0) = "ALL_BUSI_PANELS_ARRAY" Then
                            array_info = line_info(1)
                            array_info = split(array_info, "~**^")
                            ReDim Preserve ALL_BUSI_PANELS_ARRAY(budget_explain, known_busi)

                            ALL_BUSI_PANELS_ARRAY(memb_numb, known_busi)             = array_info(0)
                            ALL_BUSI_PANELS_ARRAY(panel_instance, known_busi)        = array_info(1)
                            ALL_BUSI_PANELS_ARRAY(busi_type, known_busi)             = array_info(2)
                            If array_info(3) = "CHECKED" Then ALL_BUSI_PANELS_ARRAY(estimate_only, known_busi) = checked
                            ALL_BUSI_PANELS_ARRAY(verif_explain, known_busi)         = array_info(4)
                            ALL_BUSI_PANELS_ARRAY(calc_method, known_busi)           = array_info(5)
                            ALL_BUSI_PANELS_ARRAY(info_month, known_busi)            = array_info(6)
                            ALL_BUSI_PANELS_ARRAY(mthd_date, known_busi)             = array_info(7)
                            ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, known_busi)        = array_info(8)
                            ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, known_busi)        = array_info(9)
                            ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, known_busi)      = array_info(10)
                            ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, known_busi)      = array_info(11)
                            ALL_BUSI_PANELS_ARRAY(income_ret_cash, known_busi)       = array_info(12)
                            ALL_BUSI_PANELS_ARRAY(income_pro_cash, known_busi)       = array_info(13)
                            ALL_BUSI_PANELS_ARRAY(cash_income_verif, known_busi)     = array_info(14)
                            ALL_BUSI_PANELS_ARRAY(expense_ret_cash, known_busi)      = array_info(15)
                            ALL_BUSI_PANELS_ARRAY(expense_pro_cash, known_busi)      = array_info(16)
                            ALL_BUSI_PANELS_ARRAY(cash_expense_verif, known_busi)    = array_info(17)
                            ALL_BUSI_PANELS_ARRAY(income_ret_snap, known_busi)       = array_info(18)
                            ALL_BUSI_PANELS_ARRAY(income_pro_snap, known_busi)       = array_info(19)
                            ALL_BUSI_PANELS_ARRAY(snap_income_verif, known_busi)     = array_info(20)
                            ALL_BUSI_PANELS_ARRAY(expense_ret_snap, known_busi)      = array_info(21)
                            ALL_BUSI_PANELS_ARRAY(expense_pro_snap, known_busi)      = array_info(22)
                            ALL_BUSI_PANELS_ARRAY(snap_expense_verif, known_busi)    = array_info(23)
                            If array_info(24) = "CHECKED" Then ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, known_busi) = checked
                            ALL_BUSI_PANELS_ARRAY(start_date, known_busi)            = array_info(25)
                            ALL_BUSI_PANELS_ARRAY(end_date, known_busi)              = array_info(26)
                            ALL_BUSI_PANELS_ARRAY(busi_desc, known_busi)             = array_info(27)
                            ALL_BUSI_PANELS_ARRAY(busi_structure, known_busi)        = array_info(28)
                            ALL_BUSI_PANELS_ARRAY(share_num, known_busi)             = array_info(29)
                            ALL_BUSI_PANELS_ARRAY(share_denom, known_busi)           = array_info(30)
                            ALL_BUSI_PANELS_ARRAY(partners_in_HH, known_busi)        = array_info(31)
                            ALL_BUSI_PANELS_ARRAY(exp_not_allwd, known_busi)         = array_info(32)
                            If array_info(33) = "CHECKED" Then ALL_BUSI_PANELS_ARRAY(verif_checkbox, known_busi) = checked
                            ALL_BUSI_PANELS_ARRAY(verif_added, known_busi)           = array_info(34)
                            If UCase(ALL_BUSI_PANELS_ARRAY(verif_added, known_busi)) = "TRUE" Then ALL_BUSI_PANELS_ARRAY(verif_added, known_busi) = True
                            If UCase(ALL_BUSI_PANELS_ARRAY(verif_added, known_busi)) = "FALSE" Then ALL_BUSI_PANELS_ARRAY(verif_added, known_busi) = False
                            ALL_BUSI_PANELS_ARRAY(budget_explain, known_busi)        = array_info(35)

                            known_busi = known_busi + 1
                        End If
                        ' For the_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                        '     known_busi = 0
                        '     box_one_info = ""
                        '     box_two_info = ""
                        '     box_three_info = ""
                        '     If ALL_BUSI_PANELS_ARRAY(estimate_only, the_busi) = checked Then box_one_info = "CHECKED"
                        '     If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, the_busi) = checked Then box_two_info = "CHECKED"
                        '     If ALL_BUSI_PANELS_ARRAY(verif_checkbox, the_busi) = checked Then box_three_info = "CHECKED"
                        '
                        '     objTextStream.WriteLine "ALL_BUSI_PANELS_ARRAY" & "^~^~^~^~^~^~^" &ALL_BUSI_PANELS_ARRAY(memb_numb, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(panel_instance, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_type, the_busi)&"~**^"&box_one_info&"~**^"&ALL_BUSI_PANELS_ARRAY(verif_explain, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(calc_method, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(info_month, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(mthd_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_ret_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_pro_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(cash_income_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_ret_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_pro_cash, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(cash_expense_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_ret_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(income_pro_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(snap_income_verif, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_ret_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(expense_pro_snap, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(snap_expense_verif, the_busi)&"~**^"&box_two_info&"~**^"&ALL_BUSI_PANELS_ARRAY(start_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(end_date, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_desc, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(busi_structure, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(share_num, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(share_denom, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(partners_in_HH, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(exp_not_allwd, the_busi)&"~**^"&box_three_info&"~**^"&ALL_BUSI_PANELS_ARRAY(verif_added, the_busi)&"~**^"&ALL_BUSI_PANELS_ARRAY(budget_explain, the_busi)
                        ' Next
                        If line_info(0) = "cit_id" Then cit_id = line_info(1)
                        If line_info(0) = "other_assets" Then other_assets = line_info(1)
                        If line_info(0) = "case_changes" Then case_changes = line_info(1)
                        If line_info(0) = "PREG" Then PREG = line_info(1)
                        If line_info(0) = "earned_income" Then earned_income = line_info(1)
                        If line_info(0) = "notes_on_rest" Then notes_on_rest = line_info(1)
                        If line_info(0) = "SCHL" Then SCHL = line_info(1)
                        If line_info(0) = "notes_on_jobs" Then notes_on_jobs = line_info(1)
                        If line_info(0) = "notes_on_cses" Then notes_on_cses = line_info(1)
                        If line_info(0) = "notes_on_time" Then notes_on_time = line_info(1)
                        If line_info(0) = "notes_on_sanction" Then notes_on_sanction = line_info(1)
                        If line_info(0) = "UNEA_INCOME_ARRAY" Then
                            array_info = line_info(1)
                            array_info = split(array_info, "~**^")
                            ReDim Preserve UNEA_INCOME_ARRAY(budget_notes, known_unea)

                            UNEA_INCOME_ARRAY(memb_numb, known_unea)                 = array_info(0)
                            UNEA_INCOME_ARRAY(panel_instance, known_unea)            = array_info(1)
                            UNEA_INCOME_ARRAY(UNEA_type, known_unea)                 = array_info(2)
                            UNEA_INCOME_ARRAY(UNEA_month, known_unea)                = array_info(3)
                            UNEA_INCOME_ARRAY(UNEA_verif, known_unea)                = array_info(4)
                            UNEA_INCOME_ARRAY(UNEA_prosp_amt, known_unea)            = array_info(5)
                            UNEA_INCOME_ARRAY(UNEA_retro_amt, known_unea)            = array_info(6)
                            UNEA_INCOME_ARRAY(UNEA_SNAP_amt, known_unea)             = array_info(7)
                            UNEA_INCOME_ARRAY(UNEA_pay_freq, known_unea)             = array_info(8)
                            UNEA_INCOME_ARRAY(UNEA_pic_date_calc, known_unea)        = array_info(9)
                            UNEA_INCOME_ARRAY(UNEA_UC_start_date, known_unea)        = array_info(10)
                            UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, known_unea)      = array_info(11)
                            UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, known_unea)       = array_info(12)
                            UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, known_unea)       = array_info(13)
                            UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, known_unea)        = array_info(14)
                            UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, known_unea)      = array_info(15)
                            UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, known_unea)         = array_info(16)
                            UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, known_unea)         = array_info(17)
                            UNEA_INCOME_ARRAY(UNEA_UC_notes, known_unea)             = array_info(18)
                            UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, known_unea)         = array_info(19)
                            UNEA_INCOME_ARRAY(UNEA_UC_account_balance, known_unea)   = array_info(20)
                            UNEA_INCOME_ARRAY(direct_CS_amt, known_unea)             = array_info(21)
                            UNEA_INCOME_ARRAY(disb_CS_amt, known_unea)               = array_info(22)
                            UNEA_INCOME_ARRAY(disb_CS_arrears_amt, known_unea)       = array_info(23)
                            UNEA_INCOME_ARRAY(direct_CS_notes, known_unea)           = array_info(24)
                            UNEA_INCOME_ARRAY(disb_CS_notes, known_unea)             = array_info(25)
                            UNEA_INCOME_ARRAY(disb_CS_arrears_notes, known_unea)     = array_info(26)
                            UNEA_INCOME_ARRAY(disb_CS_months, known_unea)            = array_info(27)
                            UNEA_INCOME_ARRAY(disb_CS_prosp_budg, known_unea)        = array_info(28)
                            UNEA_INCOME_ARRAY(disb_CS_arrears_months, known_unea)    = array_info(29)
                            UNEA_INCOME_ARRAY(disb_CS_arrears_budg, known_unea)      = array_info(30)
                            UNEA_INCOME_ARRAY(UNEA_RSDI_amt, known_unea)             = array_info(31)
                            UNEA_INCOME_ARRAY(UNEA_RSDI_notes, known_unea)           = array_info(32)
                            UNEA_INCOME_ARRAY(UNEA_SSI_amt, known_unea)              = array_info(33)
                            UNEA_INCOME_ARRAY(UNEA_SSI_notes, known_unea)            = array_info(34)
                            UNEA_INCOME_ARRAY(UC_exists, known_unea)                 = array_info(35)
                            If UCase(UNEA_INCOME_ARRAY(UC_exists, known_unea)) = "TRUE" Then UNEA_INCOME_ARRAY(UC_exists, known_unea) = True
                            If UCase(UNEA_INCOME_ARRAY(UC_exists, known_unea)) = "FALSE" Then UNEA_INCOME_ARRAY(UC_exists, known_unea) = False
                            UNEA_INCOME_ARRAY(CS_exists, known_unea)                 = array_info(36)
                            If UCase(UNEA_INCOME_ARRAY(CS_exists, known_unea)) = "TRUE" Then UNEA_INCOME_ARRAY(CS_exists, known_unea) = True
                            If UCase(UNEA_INCOME_ARRAY(CS_exists, known_unea)) = "FALSE" Then UNEA_INCOME_ARRAY(CS_exists, known_unea) = False
                            UNEA_INCOME_ARRAY(SSA_exists, known_unea)                = array_info(37)
                            If UCase(UNEA_INCOME_ARRAY(SSA_exists, known_unea)) = "TRUE" Then UNEA_INCOME_ARRAY(SSA_exists, known_unea) = True
                            If UCase(UNEA_INCOME_ARRAY(SSA_exists, known_unea)) = "FALSE" Then UNEA_INCOME_ARRAY(SSA_exists, known_unea) = False
                            UNEA_INCOME_ARRAY(calc_button, known_unea)               = array_info(38)
                            UNEA_INCOME_ARRAY(budget_notes, known_unea)              = array_info(39)

                            known_unea = known_unea + 1
                        End If

                        ' UNEA_INCOME_ARRAY(memb_numb, the_unea)0
                        ' UNEA_INCOME_ARRAY(panel_instance, the_unea)1
                        ' UNEA_INCOME_ARRAY(UNEA_type, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_month, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_verif, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_prosp_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_retro_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_SNAP_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_pay_freq, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_pic_date_calc, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_start_date, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_notes, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_UC_account_balance, the_unea)
                        ' UNEA_INCOME_ARRAY(direct_CS_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_arrears_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(direct_CS_notes, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_notes, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_arrears_notes, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_months, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_prosp_budg, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_arrears_months, the_unea)
                        ' UNEA_INCOME_ARRAY(disb_CS_arrears_budg, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_RSDI_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_RSDI_notes, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_SSI_amt, the_unea)
                        ' UNEA_INCOME_ARRAY(UNEA_SSI_notes, the_unea)
                        ' UNEA_INCOME_ARRAY(UC_exists, the_unea)
                        ' UNEA_INCOME_ARRAY(CS_exists, the_unea)
                        ' UNEA_INCOME_ARRAY(SSA_exists, the_unea)
                        ' UNEA_INCOME_ARRAY(calc_button, the_unea)
                        ' UNEA_INCOME_ARRAY(budget_notes, the_unea)

                        ' For the_unea = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                        '     known_unea = 0
                        '     objTextStream.WriteLine "UNEA_INCOME_ARRAY" & "^~^~^~^~^~^~^" &UNEA_INCOME_ARRAY(memb_numb, the_unea)&"~**^"&UNEA_INCOME_ARRAY(panel_instance, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_type, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_month, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_verif, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_prosp_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_retro_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SNAP_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_pay_freq, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_pic_date_calc, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_start_date, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_UC_account_balance, the_unea)&"~**^"&UNEA_INCOME_ARRAY(direct_CS_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(direct_CS_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_months, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_prosp_budg, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_months, the_unea)&"~**^"&UNEA_INCOME_ARRAY(disb_CS_arrears_budg, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_RSDI_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_RSDI_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SSI_amt, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UNEA_SSI_notes, the_unea)&"~**^"&UNEA_INCOME_ARRAY(UC_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(CS_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(SSA_exists, the_unea)&"~**^"&UNEA_INCOME_ARRAY(calc_button, the_unea)&"~**^"&UNEA_INCOME_ARRAY(budget_notes, the_unea)
                        ' Next
                        If line_info(0) = "notes_on_wreg" Then notes_on_wreg = line_info(1)
                        If line_info(0) = "full_abawd_info" Then full_abawd_info = line_info(1)
                        If line_info(0) = "notes_on_abawd" Then notes_on_abawd = line_info(1)
                        If line_info(0) = "notes_on_abawd_two" Then notes_on_abawd_two = line_info(1)
                        If line_info(0) = "notes_on_abawd_three" Then notes_on_abawd_three = line_info(1)
                        If line_info(0) = "programs_applied_for" Then programs_applied_for = line_info(1)
                        ' If TIKL_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "TIKL_checkbox" and line_info(1) = "CHECKED" Then TIKL_checkbox = checked
                        If line_info(0) = "interview_memb_list" Then interview_memb_list = line_info(1)
                        If line_info(0) = "shel_memb_list" Then shel_memb_list = line_info(1)
                        If line_info(0) = "verification_memb_list" Then verification_memb_list = line_info(1)
                        If line_info(0) = "notes_on_busi" Then notes_on_busi = line_info(1)
                        'DLG 1
                        ' If Used_Interpreter_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "Used_Interpreter_checkbox" and line_info(1) = "CHECKED" Then Used_Interpreter_checkbox = checked
                        ' If line_info(0) = "how_app_rcvd" Then how_app_rcvd = line_info(1)
                        If line_info(0) = "arep_id_info" Then arep_id_info = line_info(1)
                        If line_info(0) = "CS_forms_sent_date" Then CS_forms_sent_date = line_info(1)
                        ' If line_info(0) = "case_changes" Then case_changes = line_info(1)
                        'DLG 5'
                        If line_info(0) = "notes_on_ssa_income" Then notes_on_ssa_income = line_info(1)
                        If line_info(0) = "notes_on_VA_income" Then notes_on_VA_income = line_info(1)
                        If line_info(0) = "notes_on_WC_income" Then notes_on_WC_income = line_info(1)
                        If line_info(0) = "other_uc_income_notes" Then other_uc_income_notes = line_info(1)
                        If line_info(0) = "notes_on_other_UNEA" Then notes_on_other_UNEA = line_info(1)

                        If line_info(0) = "hest_information" Then hest_information = line_info(1)
                        ' If line_info(0) = "notes_on_acut" Then notes_on_acut = line_info(1)
                        ' If line_info(0) = "notes_on_coex" Then notes_on_coex = line_info(1)
                        ' If line_info(0) = "notes_on_dcex" Then notes_on_dcex = line_info(1)
                        If line_info(0) = "notes_on_other_deduction" Then notes_on_other_deduction = line_info(1)
                        If line_info(0) = "expense_notes" Then expense_notes = line_info(1)
                        ' If address_confirmation_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "address_confirmation_checkbox" and line_info(1) = "CHECKED" Then address_confirmation_checkbox = checked
                        If line_info(0) = "manual_total_shelter" Then manual_total_shelter = line_info(1)
                        If line_info(0) = "manual_amount_used" Then manual_amount_used = line_info(1)
                        If UCase(manual_amount_used) = "TRUE" Then manual_amount_used = True
                        If UCase(manual_amount_used) = "FALSE" Then manual_amount_used = False
                        If line_info(0) = "app_month_assets" Then app_month_assets = line_info(1)
                        ' If confirm_no_account_panel_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "confirm_no_account_panel_checkbox" and line_info(1) = "CHECKED" Then confirm_no_account_panel_checkbox = checked
                        If line_info(0) = "notes_on_other_assets" Then notes_on_other_assets = line_info(1)
                        If line_info(0) = "MEDI" Then MEDI = line_info(1)
                        If line_info(0) = "DISQ" Then DISQ = line_info(1)
                        'EXP DET'
                        If line_info(0) = "full_determination_done" Then full_determination_done = line_info(1)
                        If UCase(full_determination_done) = "TRUE" Then full_determination_done = True
                        If UCase(full_determination_done) = "FALSE" Then full_determination_done = False
                        ' Call run_expedited_determination_script_functionality(
                        ' If line_info(0) = "xfs_screening" Then xfs_screening = line_info(1)
                        ' If line_info(0) = "caf_one_income" Then caf_one_income = line_info(1)
                        ' If line_info(0) = "caf_one_assets" Then caf_one_assets = line_info(1)
                        ' If line_info(0) = "caf_one_rent" Then caf_one_rent = line_info(1)
                        ' If line_info(0) = "caf_one_utilities" Then caf_one_utilities = line_info(1)
                        If line_info(0) = "determined_income" Then determined_income = line_info(1)
                        If line_info(0) = "determined_assets" Then determined_assets = line_info(1)
                        If line_info(0) = "determined_shel" Then determined_shel = line_info(1)
                        If line_info(0) = "determined_utilities" Then determined_utilities = line_info(1)
                        If line_info(0) = "calculated_resources" Then calculated_resources = line_info(1)
                        If line_info(0) = "calculated_expenses" Then calculated_expenses = line_info(1)
                        If line_info(0) = "calculated_low_income_asset_test" Then calculated_low_income_asset_test = line_info(1)
                        If UCase(calculated_low_income_asset_test) = "TRUE" Then calculated_low_income_asset_test = True
                        If UCase(calculated_low_income_asset_test) = "FALSE" Then calculated_low_income_asset_test = False
                        If line_info(0) = "calculated_resources_less_than_expenses_test" Then calculated_resources_less_than_expenses_test = line_info(1)
                        If UCase(calculated_resources_less_than_expenses_test) = "TRUE" Then calculated_resources_less_than_expenses_test = True
                        If UCase(calculated_resources_less_than_expenses_test) = "FALSE" Then calculated_resources_less_than_expenses_test = False
                        If line_info(0) = "is_elig_XFS" Then is_elig_XFS = line_info(1)
                        If UCase(is_elig_XFS) = "TRUE" Then is_elig_XFS = True
                        If UCase(is_elig_XFS) = "FALSE" Then is_elig_XFS = False
                        If line_info(0) = "approval_date" Then approval_date = line_info(1)
                        If line_info(0) = "applicant_id_on_file_yn" Then applicant_id_on_file_yn = line_info(1)
                        If line_info(0) = "applicant_id_through_SOLQ" Then applicant_id_through_SOLQ = line_info(1)
                        If line_info(0) = "delay_explanation" Then delay_explanation = line_info(1)
                        ' If line_info(0) = "snap_denial_date" Then snap_denial_date = line_info(1)
                        If line_info(0) = "snap_denial_explain" Then snap_denial_explain = line_info(1)
                        If line_info(0) = "case_assesment_text" Then case_assesment_text = line_info(1)
                        If line_info(0) = "next_steps_one" Then next_steps_one = line_info(1)
                        If line_info(0) = "next_steps_two" Then next_steps_two = line_info(1)
                        If line_info(0) = "next_steps_three" Then next_steps_three = line_info(1)
                        If line_info(0) = "next_steps_four" Then next_steps_four = line_info(1)
                        If line_info(0) = "postponed_verifs_yn" Then postponed_verifs_yn = line_info(1)
                        If line_info(0) = "list_postponed_verifs" Then list_postponed_verifs = line_info(1)
                        If line_info(0) = "day_30_from_application" Then day_30_from_application = line_info(1)
                        If line_info(0) = "other_snap_state" Then other_snap_state = line_info(1)
                        If line_info(0) = "other_state_reported_benefit_end_date" Then other_state_reported_benefit_end_date = line_info(1)
                        If line_info(0) = "other_state_benefits_openended" Then other_state_benefits_openended = line_info(1)
                        If UCase(other_state_benefits_openended) = "TRUE" Then other_state_benefits_openended = True
                        If UCase(other_state_benefits_openended) = "FALSE" Then other_state_benefits_openended = False
                        If line_info(0) = "other_state_contact_yn" Then other_state_contact_yn = line_info(1)
                        If line_info(0) = "other_state_verified_benefit_end_date" Then other_state_verified_benefit_end_date = line_info(1)
                        If line_info(0) = "mn_elig_begin_date" Then mn_elig_begin_date = line_info(1)
                        If line_info(0) = "action_due_to_out_of_state_benefits" Then action_due_to_out_of_state_benefits = line_info(1)
                        If line_info(0) = "case_has_previously_postponed_verifs_that_prevent_exp_snap" Then case_has_previously_postponed_verifs_that_prevent_exp_snap = line_info(1)
                        If UCase(case_has_previously_postponed_verifs_that_prevent_exp_snap) = "TRUE" Then case_has_previously_postponed_verifs_that_prevent_exp_snap = True
                        If UCase(case_has_previously_postponed_verifs_that_prevent_exp_snap) = "FALSE" Then case_has_previously_postponed_verifs_that_prevent_exp_snap = False
                        If line_info(0) = "prev_post_verif_assessment_done" Then prev_post_verif_assessment_done = line_info(1)
                        If UCase(prev_post_verif_assessment_done) = "TRUE" Then prev_post_verif_assessment_done = True
                        If UCase(prev_post_verif_assessment_done) = "FALSE" Then prev_post_verif_assessment_done = False
                        If line_info(0) = "previous_date_of_application" Then previous_date_of_application = line_info(1)
                        If line_info(0) = "previous_expedited_package" Then previous_expedited_package = line_info(1)
                        If line_info(0) = "prev_verifs_mandatory_yn" Then prev_verifs_mandatory_yn = line_info(1)
                        If line_info(0) = "prev_verif_list" Then prev_verif_list = line_info(1)
                        If line_info(0) = "curr_verifs_postponed_yn" Then curr_verifs_postponed_yn = line_info(1)
                        If line_info(0) = "ongoing_snap_approved_yn" Then ongoing_snap_approved_yn = line_info(1)
                        If line_info(0) = "prev_post_verifs_recvd_yn" Then prev_post_verifs_recvd_yn = line_info(1)
                        If line_info(0) = "delay_action_due_to_faci" Then delay_action_due_to_faci = line_info(1)
                        If UCase(delay_action_due_to_faci) = "TRUE" Then delay_action_due_to_faci = True
                        If UCase(delay_action_due_to_faci) = "FALSE" Then delay_action_due_to_faci = False
                        If line_info(0) = "deny_snap_due_to_faci" Then deny_snap_due_to_faci = line_info(1)
                        If UCase(deny_snap_due_to_faci) = "TRUE" Then deny_snap_due_to_faci = True
                        If UCase(deny_snap_due_to_faci) = "FALSE" Then deny_snap_due_to_faci = False
                        If line_info(0) = "faci_review_completed" Then faci_review_completed = line_info(1)
                        If UCase(faci_review_completed) = "TRUE" Then faci_review_completed = True
                        If UCase(faci_review_completed) = "FALSE" Then faci_review_completed = False
                        If line_info(0) = "facility_name" Then facility_name = line_info(1)
                        If line_info(0) = "snap_inelig_faci_yn" Then snap_inelig_faci_yn = line_info(1)
                        If line_info(0) = "faci_entry_date" Then faci_entry_date = line_info(1)
                        If line_info(0) = "faci_release_date" Then faci_release_date = line_info(1)
                        If line_info(0) = "release_date_unknown_checkbox" AND line_info(1) = "CHECKED" Then release_date_unknown_checkbox = checked
                        If line_info(0) = "release_within_30_days_yn" Then release_within_30_days_yn = line_info(1)

                        If line_info(0) = "next_er_month" Then next_er_month = line_info(1)
                        If line_info(0) = "next_er_year" Then next_er_year = line_info(1)
                        If line_info(0) = "CAF_status" Then CAF_status = line_info(1)
                        If line_info(0) = "actions_taken" Then actions_taken = line_info(1)
                        ' If application_signed_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "application_signed_checkbox" and line_info(1) = "CHECKED" Then application_signed_checkbox = checked
                        ' If eDRS_sent_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "eDRS_sent_checkbox" and line_info(1) = "CHECKED" Then eDRS_sent_checkbox = checked
                        ' If updated_MMIS_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "updated_MMIS_checkbox" and line_info(1) = "CHECKED" Then updated_MMIS_checkbox = checked
                        ' If WF1_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "WF1_checkbox" and line_info(1) = "CHECKED" Then WF1_checkbox = checked
                        ' If Sent_arep_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "Sent_arep_checkbox" and line_info(1) = "CHECKED" Then Sent_arep_checkbox = checked
                        ' If intake_packet_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "intake_packet_checkbox" and line_info(1) = "CHECKED" Then intake_packet_checkbox = checked
                        ' If IAA_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "IAA_checkbox" and line_info(1) = "CHECKED" Then IAA_checkbox = checked
                        ' If recert_period_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "recert_period_checkbox" and line_info(1) = "CHECKED" Then recert_period_checkbox = checked
                        ' If R_R_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "R_R_checkbox" and line_info(1) = "CHECKED" Then R_R_checkbox = checked
                        ' If E_and_T_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "E_and_T_checkbox" and line_info(1) = "CHECKED" Then E_and_T_checkbox = checked
                        ' If elig_req_explained_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "elig_req_explained_checkbox" and line_info(1) = "CHECKED" Then elig_req_explained_checkbox = checked
                        ' If benefit_payment_explained_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "benefit_payment_explained_checkbox" and line_info(1) = "CHECKED" Then benefit_payment_explained_checkbox = checked
                        If line_info(0) = "other_notes" Then other_notes = line_info(1)
                        ' If client_delay_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "client_delay_checkbox" and line_info(1) = "CHECKED" Then client_delay_checkbox = checked
                        ' If TIKL_checkbox = checked Then objTextStream.WriteLine
                        ' If line_info(0) = "TIKL_checkbox" and line_info(1) = "CHECKED" Then TIKL_checkbox = checked
                        ' If client_delay_TIKL_checkbox = checked Then objTextStream.WriteLine
                        If line_info(0) = "client_delay_TIKL_checkbox" and line_info(1) = "CHECKED" Then client_delay_TIKL_checkbox = checked
                        If line_info(0) = "verif_req_form_sent_date" Then verif_req_form_sent_date = line_info(1)
                        If line_info(0) = "worker_signature" Then worker_signature = line_info(1)



                    End If
                Next
            End If
        End If
    End With
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
					  Text 20, 205, 205, 20, "SNAP should be APPROVEED "
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
	If is_elig_XFS = True Then email_body = email_body & "Case appears EXPEDITED." & vbCr
	If is_elig_XFS = False Then email_body = email_body & "Case does NOT appear Expedtied." & vbCr
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

	email_body = "~~This email is generated from wihtin the 'Expedited Determination' Script.~~" & vbCr & vbCr & email_body
	call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", True)
	' call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", False)
	' create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
end function

'FUNCTIONS =================================================================================================================
'This function will message box the err_msg from this script, outlining it by dialog and adding the headers.
'This function should be added to the end of the dialogs after the review button and at the end of dialog 8 after the error message collection.
function display_errors(the_err_msg, execute_nav)
    If the_err_msg <> "" Then       'If the error message is blank - there is nothing to show.
        If left(the_err_msg, 3) = "~!~" Then the_err_msg = right(the_err_msg, len(the_err_msg) - 3)     'Trimming the message so we don't have a blank array item
        err_array = split(the_err_msg, "~!~")           'making the list of errors an array.

        error_message = ""                              'blanking out variables
        msg_header = ""
        for each message in err_array                   'going through each error message to order them and add headers'
            current_listing = left(message, 1)          'This is the dialog the error came from
            If current_listing <> msg_header Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
                If current_listing = "1" Then tagline = ": Personal Information"        'Adding a specific tagline to the header for the errors
                If current_listing = "2" Then tagline = ": JOBS"
                If current_listing = "3" Then tagline = ": BUSI"
                If current_listing = "4" Then tagline = ": Child Support"
                If current_listing = "5" Then tagline = ": Unearned Income"
                If current_listing = "6" Then tagline = ": WREG, Expenses, Address"
                If current_listing = "7" Then tagline = ": Assets and Misc."
                If current_listing = "8" Then tagline = ": Interview Detail"
                error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
            End If
            if msg_header = "" Then back_to_dialog = current_listing
            msg_header = current_listing        'setting for the next loop

            message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

            error_message = error_message & vbNewLine & right(message, len(message) - 2)        'Adding the error information to the message list.
        Next

        'This is the display of all of the messages.
        view_errors = MsgBox("In order to complete the script and CASE/NOTE, additional details need to be added or refined. Please review and update." & vbNewLine & error_message, vbCritical, "Review detail required in Dialogs")

        'The function can be operated without moving to a different dialog or not. The only time this will be activated is at the end of dialog 8.
        If execute_nav = TRUE Then
            If back_to_dialog = "1" Then ButtonPressed = dlg_one_button         'This calls another function to go to the first dialog that had an error
            If back_to_dialog = "2" Then ButtonPressed = dlg_two_button
            If back_to_dialog = "3" Then ButtonPressed = dlg_three_button
            If back_to_dialog = "4" Then ButtonPressed = dlg_four_button
            If back_to_dialog = "5" Then ButtonPressed = dlg_five_button
            If back_to_dialog = "6" Then ButtonPressed = dlg_six_button
            If back_to_dialog = "7" Then ButtonPressed = dlg_seven_button
            If back_to_dialog = "8" Then ButtonPressed = dlg_eight_button

            Call assess_button_pressed          'this is where the navigation happens
        End If
    End If
End Function

Function HCRE_panel_bypass()
    'handling for cases that do not have a completed HCRE panel
    PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
    Do
        EMReadscreen HCRE_panel_check, 4, 2, 50
        If HCRE_panel_check = "HCRE" then
            PF10	'exists edit mode in cases where HCRE isn't complete for a member
            PF3
        END IF
    Loop until HCRE_panel_check <> "HCRE"
End Function

'This function calls the dialog to determine and assess the household Composition
'This also determines the members that are including in gathering information.
function HH_comp_dialog(HH_member_array)
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
    EMWriteScreen "01", 20, 76
    transmit

	EMReadScreen id_ver_code, 2, 9, 68
	If id_ver_code <> "__" AND id_ver_code <> "NO" Then applicant_id_on_file_yn = "Yes"
	If id_ver_code = "__" OR id_ver_code = "NO" Then applicant_id_on_file_yn = "No"

    member_count = 0            'resetting these counts/variables
    adult_cash_count = 0
    child_cash_count = 0
    adult_snap_count = 0
    child_snap_count = 0
    adult_emer_count = 0
    child_emer_count = 0
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 2, 4, 33
        EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
        If access_denied_check = "ACCESS DENIED" Then
            PF10
            EMWaitReady 0, 0
            last_name = "UNABLE TO FIND"
            first_name = " - Access Denied"
            mid_initial = ""
        Else
    		EMReadscreen last_name, 25, 6, 30
    		EMReadscreen first_name, 12, 6, 63
    		EMReadscreen mid_initial, 1, 6, 79
            EMReadScreen memb_age, 3, 8, 76
            memb_age = trim(memb_age)
            If memb_age = "" Then memb_age = 0
            memb_age = memb_age * 1
    		last_name = trim(replace(last_name, "_", ""))
    		first_name = trim(replace(first_name, "_", ""))
    		mid_initial = replace(mid_initial, "_", "")
            EMReadScreen id_verif_code, 2, 9, 68


            EMReadScreen rel_to_applcnt, 2, 10, 42              'reading the relationship from MEMB'
            If rel_to_applcnt = "02" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Spouse of Memb 01.; "
            If rel_to_applcnt = "04" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Parent of Memb 01.; "
            If rel_to_applcnt = "05" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Sibling of Memb 01.; "
            If rel_to_applcnt = "12" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Niece of Memb 01.; "
            If rel_to_applcnt = "13" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Nephew of Memb 01.; "
            If rel_to_applcnt = "15" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Grandparent of Memb 01.; "
            If rel_to_applcnt = "16" Then relationship_detail = relationship_detail & "Memb " & ref_nbr & " is the Grandchild of Memb 01.; "
        End If

        ReDim Preserve ALL_MEMBERS_ARRAY(clt_notes, member_count)       'resizing the array to add the next household member

        ALL_MEMBERS_ARRAY(memb_numb, member_count) = ref_nbr            'adding client information to the array
        ALL_MEMBERS_ARRAY(clt_name, member_count) = last_name & ", " & first_name & " " & mid_initial
        ALL_MEMBERS_ARRAY(full_clt, member_count) = ref_nbr & " - " & first_name & " " & last_name
        ALL_MEMBERS_ARRAY(clt_age, member_count) = memb_age

        If id_verif_code = "BC" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "BC - Birth Certificate"
        If id_verif_code = "RE" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "RE - Religious Record"
        If id_verif_code = "DL" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "DL - Drivers License/ST ID"
        If id_verif_code = "DV" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "DV - Divorce Decree"
        If id_verif_code = "AL" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "AL - Alien Card"
        If id_verif_code = "AD" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "AD - Arrival//Depart"
        If id_verif_code = "DR" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "DR - Doctor Stmt"
        If id_verif_code = "PV" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "PV - Passport/Visa"
        If id_verif_code = "OT" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "OT - Other Document"
        If id_verif_code = "NO" Then ALL_MEMBERS_ARRAY(clt_id_verif, member_count) = "NO - No Verif Prvd"

        If cash_checkbox = checked or GRH_checkbox = checked Then             'If Cash is selected
            If cash_checkbox = unchecked Then
				If member_count = 0 Then
					ALL_MEMBERS_ARRAY(include_cash_checkbox, member_count) = checked    'default to having the counted boxes checked for SNAP
					ALL_MEMBERS_ARRAY(count_cash_checkbox, member_count) = checked
					If memb_age > 18 then       'Adding to the cash count
						adult_cash_count = adult_cash_count + 1
					Else
						child_cash_count = child_cash_count + 1
					End If
				End If
			Else
				ALL_MEMBERS_ARRAY(include_cash_checkbox, member_count) = checked    'default to having the counted boxes checked for SNAP
				ALL_MEMBERS_ARRAY(count_cash_checkbox, member_count) = checked
				If memb_age > 18 then       'Adding to the cash count
					adult_cash_count = adult_cash_count + 1
				Else
					child_cash_count = child_cash_count + 1
				End If
			End If
        End If
        If SNAP_checkbox = checked Then             'If SNAP is selected
            ALL_MEMBERS_ARRAY(include_snap_checkbox, member_count) = checked    'default to having the counted boxes checked for SNAP
            ALL_MEMBERS_ARRAY(count_snap_checkbox, member_count) = checked
            If memb_age > 21 then       'adding to the snap household member count
                adult_snap_count = adult_snap_count + 1
            Else
                child_snap_count = child_snap_count + 1
            End If
        End If
        If EMER_checkbox = checked Then             'If EMER is selected
            ALL_MEMBERS_ARRAY(include_emer_checkbox, member_count) = checked    'default to having the counted boxes checked for EMER
            ALL_MEMBERS_ARRAY(count_emer_checkbox, member_count) = checked
            If memb_age > 18 then       'Adding to the EMER count
                adult_emer_count = adult_emer_count + 1
            Else
                child_emer_count = child_emer_count + 1
            End If
        End If

		client_string = ref_nbr & last_name & first_name & mid_initial            'creating an array of all of the clients
		client_array = client_array & client_string & "|"
		transmit      'Going to the next MEMB panel
		Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
        member_count = member_count + 1
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

    Call navigate_to_MAXIS_screen("STAT", "PARE")       'Going to get relationship information from PARE
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2) 'looping through each member
        EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76     'Going to PARE for each member
        transmit

        EMReadScreen panel_check, 14, 24, 13        'Making sure there is a PARE panel to read from
        If panel_check <> "DOES NOT EXIST" Then
            pare_row = 8                            'start of information on PARE
            Do
                EMReadScreen child_ref_nbr, 2, pare_row, 24     'Reading child, relationship and verif
                EMReadScreen rela_type, 1, pare_row, 53
                EMReadScreen rela_verif, 2, pare_row, 71
                If child_ref_nbr = "__" then exit do

                If rela_type = "1" then relationship_type = "Parent"            'Changing the code for the relationship to the words are used instead of code.
                If rela_type = "2" then relationship_type = "Stepparent"
                If rela_type = "3" then relationship_type = "Grandparent"
                If rela_type = "4" then relationship_type = "Relative Caregiver"
                If rela_type = "5" then relationship_type = "Foster parent"
                If rela_type = "6" then relationship_type = "Caregiver"
                If rela_type = "7" then relationship_type = "Guardian"
                If rela_type = "8" then relationship_type = "Relative"

                If rela_verif = "BC" Then relationship_verif = "Birth Certificate"      'Change the code for verif to full words for readability
                If rela_verif = "AR" Then relationship_verif = "Adoption Records"
                If rela_verif = "LG" Then relationship_verif = "Legal Guardian"
                If rela_verif = "RE" Then relationship_verif = "Religious Records"
                If rela_verif = "HR" Then relationship_verif = "Hospital Records"
                If rela_verif = "RP" Then relationship_verif = "Recognition of Parantage"
                If rela_verif = "OT" Then relationship_verif = "Other"
                If rela_verif = "NO" Then relationship_verif = "NONE"

                'Here is where the relationship information is added to the field of the dialog
                If child_ref_nbr <> "__" Then relationship_detail = relationship_detail & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " is the " & relationship_type & " of Memb " & child_ref_nbr & " - Verif: " & relationship_verif & "; "
                pare_row = pare_row + 1 'going to the next rwo
                If pare_row = 18 then
                    PF20 'shift PF8
                    EmReadscreen last_screen, 21, 24, 2
                    If last_screen = "THIS IS THE LAST PAGE" then
                        exit do
                    Else
                        pare_row = 8
                    End if
                End if
            Loop
        End If
    Next

    client_array = TRIM(client_array)
    client_array = split(client_array, "|")
    If SNAP_checkbox = checked then call read_EATS_panel        'If SNAP, we need to read EATS. This is a local function.

    Do
        Do
            err_msg = ""
            adult_cash_count = adult_cash_count & ""            'Setting variables to be strings
            child_cash_count = child_cash_count & ""
            adult_snap_count = adult_snap_count & ""
            child_snap_count = child_snap_count & ""
            adult_emer_count = adult_emer_count & ""
            child_emer_count = child_emer_count & ""

            'Dialog of the Household Composition
            dlg_len = 115 + (15 * UBound(ALL_MEMBERS_ARRAY, 2))     'setting the size of the dialog based on the number of household members
            if dlg_len < 145 Then dlg_len = 145                     'This is the minimum height of the dialog
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 446, dlg_len, "HH Composition Dialog"
              Text 10, 10, 250, 10, "This dialog will clarify the household relationships and details for the case."
              Text 105, 25, 100, 10, "Included and Counted in Grant"
              x_pos = 110
              count_cash_pos = x_pos + 5
              If cash_checkbox = checked or GRH_checkbox = checked Then
                Text x_pos, 40, 20, 10, "Cash"
                x_pos = x_pos + 35
              End If
              count_snap_pos = x_pos + 5
              If SNAP_checkbox = checked Then
                Text x_pos, 40, 20, 10, "SNAP"
                x_pos = x_pos + 35
              End If
              count_emer_pos = x_pos + 5
              If EMER_checkbox = checked Then Text x_pos, 40, 20, 10, "EMER"
              Text 230, 25, 90, 10, "Income Counted - Deeming"
              x_pos = 230
              deem_cash_pos = x_pos + 5
              If cash_checkbox = checked or GRH_checkbox = checked Then
                Text x_pos, 40, 20, 10, "Cash"
                x_pos = x_pos + 35
              End If
              deem_snap_pos = x_pos + 5
              If SNAP_checkbox = checked Then
                Text x_pos, 40, 20, 10, "SNAP"
                x_pos = x_pos + 35
              End If
              deem_emer_pos = x_pos + 5
              If EMER_checkbox = checked Then Text x_pos, 40, 20, 10, "EMER"
              GroupBox 330, 5, 105, 120, "HH Count by program"
              Text 335, 15, 100, 20, "Enter the number of adults and children for each program"
              Text 370, 35, 20, 10, "Adults"
              Text 400, 35, 30, 10, "Children"
              hh_comp_pos = 45
              If cash_checkbox = checked or GRH_checkbox = checked Then
                  Text 345, hh_comp_pos + 5, 20, 10, "Cash"
                  EditBox 370, hh_comp_pos, 20, 15, adult_cash_count
                  EditBox 405, hh_comp_pos, 20, 15, child_cash_count
                  CheckBox 355, hh_comp_pos + 20, 75, 10, "Pregnant Caregiver", pregnant_caregiver_checkbox
                  hh_comp_pos = hh_comp_pos + 35
              End If
              If SNAP_checkbox = checked Then
                  Text 345, hh_comp_pos + 5, 20, 10, "SNAP"
                  EditBox 370, hh_comp_pos, 20, 15, adult_snap_count
                  EditBox 405, hh_comp_pos, 20, 15, child_snap_count
                  hh_comp_pos = hh_comp_pos + 20
              End If
              If EMER_checkbox = checked then
                  Text 345, hh_comp_pos, 25, 10, "EMER"
                  EditBox 370, hh_comp_pos, 20, 15, adult_emer_count
                  EditBox 405, hh_comp_pos, 20, 15, child_emer_count
              End If
              y_pos = 55
              For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                  Text 10, y_pos, 100, 10, ALL_MEMBERS_ARRAY(clt_name, each_member)
                  If cash_checkbox = checked or GRH_checkbox = checked Then CheckBox count_cash_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member)
                  If SNAP_checkbox = checked Then CheckBox count_snap_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member)
                  If EMER_checkbox = checked then CheckBox count_emer_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member)
                  If cash_checkbox = checked or GRH_checkbox = checked Then CheckBox deem_cash_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member)
                  If SNAP_checkbox = checked Then CheckBox deem_snap_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member)
                  If EMER_checkbox = checked then CheckBox deem_emer_pos, y_pos, 10, 10, "", ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member)
                  y_pos = y_pos + 15
              Next
              if y_pos < 100 Then Y_pos = 100
              y_pos = y_pos + 5
              Text 10, y_pos + 5, 25, 10, "EATS:"
              EditBox 35, y_pos, 290, 15, EATS
              Text 10, y_pos + 25, 90, 10, "Household Relationships:"
              EditBox 105, y_pos + 20, 220, 15, relationship_detail
              ButtonGroup ButtonPressed
                OkButton 335, y_pos + 20, 50, 15
                CancelButton 390, y_pos + 20, 50, 15
            EndDialog

            dialog Dialog1
            cancel_without_confirmation

            If trim(adult_cash_count) = "" Then adult_cash_count = 0            ''
            If trim(child_cash_count) = "" Then child_cash_count = 0
            If IsNumeric(adult_cash_count) = False and IsNumeric(child_cash_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the Cash program."

            If trim(adult_snap_count) = "" Then adult_snap_count = 0
            If trim(child_snap_count) = "" Then child_snap_count = 0
            If IsNumeric(adult_snap_count) = False and IsNumeric(child_snap_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the SNAP program."
            If SNAP_checkbox = checked AND trim(EATS) = "" Then err_msg = err_msg & vbNewLine & "* Clarify who purchases and prepares together since SNAP is being considered."

            If trim(adult_emer_count) = "" Then adult_emer_count = 0
            If trim(child_emer_count) = "" Then child_emer_count = 0
            If IsNumeric(adult_emer_count) = False and IsNumeric(child_emer_count) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid count for the number of adults and children for the EMER program."

            adult_cash_count = adult_cash_count * 1
            child_cash_count = child_cash_count * 1
            adult_snap_count = adult_snap_count * 1
            child_snap_count = child_snap_count * 1
            adult_emer_count = adult_emer_count * 1
            child_emer_count = child_emer_count * 1

            If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    HH_member_count = 0

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        ALL_MEMBERS_ARRAY(gather_detail, each_member) = FALSE
        If ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member) = checked Then
            ReDim Preserve HH_member_array(HH_member_count)
            HH_member_array(HH_member_count) = ALL_MEMBERS_ARRAY(memb_numb, each_member)
            HH_member_count = HH_member_count + 1
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            ReDim Preserve HH_member_array(HH_member_count)
            HH_member_array(HH_member_count) = ALL_MEMBERS_ARRAY(memb_numb, each_member)
            HH_member_count = HH_member_count + 1
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member) = checked Then
            ReDim Preserve HH_member_array(HH_member_count)
            HH_member_array(HH_member_count) = ALL_MEMBERS_ARRAY(memb_numb, each_member)
            HH_member_count = HH_member_count + 1
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member) = checked Then
            ReDim Preserve HH_member_array(HH_member_count)
            HH_member_array(HH_member_count) = ALL_MEMBERS_ARRAY(memb_numb, each_member)
            HH_member_count = HH_member_count + 1
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member) = checked Then
            ReDim Preserve HH_member_array(HH_member_count)
            HH_member_array(HH_member_count) = ALL_MEMBERS_ARRAY(memb_numb, each_member)
            HH_member_count = HH_member_count + 1
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        ElseIf ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member) = checked Then
            ReDim Preserve HH_member_array(HH_member_count)
            HH_member_array(HH_member_count) = ALL_MEMBERS_ARRAY(memb_numb, each_member)
            HH_member_count = HH_member_count + 1
            ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE
        End If
    Next

	' HH_member_list = TRIM(HH_member_list)							'Cleaning up array for ease of use.
    ' HH_member_list = REPLACE(HH_member_list, "  ", " ")
	' HH_member_array = SPLIT(HH_member_list, " ")
    ' MsgBox "All members ubound - " & UBound(ALL_MEMBERS_ARRAY, 2)
end function

function read_BUSI_panel()
    EMReadScreen income_type, 2, 5, 37
    EMReadScreen retro_rpt_hrs, 3, 13, 60
    EMReadScreen prosp_rpt_hrs, 3, 13, 74
    EMReadScreen retro_min_wg_hrs, 3, 14, 60
    EMReadScreen prosp_min_wg_hrs, 3, 14, 74
    EMReadScreen self_emp_method, 2, 16, 53
    EMReadScreen method_date, 8, 16, 63
    EMReadScreen income_start, 8, 5, 55
    EMReadScreen income_end, 8, 5, 72

    If income_type = "01" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "01 - Farming"
    If income_type = "02" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "02 - Real Estate"
    If income_type = "03" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "03 - Home Product Sales"
    If income_type = "04" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "04 - Other Sales"
    If income_type = "05" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "05 - Personal Services"
    If income_type = "06" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "06 - Paper Route"
    If income_type = "07" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "07 - In Home Daycare"
    If income_type = "08" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "08 - Rental Income"
    If income_type = "09" Then ALL_BUSI_PANELS_ARRAY(busi_type, busi_count) = "09 - Other"
    ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, busi_count) = trim(retro_rpt_hrs)
    ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, busi_count) = trim(prosp_rpt_hrs)
    ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, busi_count) = trim(retro_min_wg_hrs)
    ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, busi_count) = trim(prosp_min_wg_hrs)
    If self_emp_method = "01" Then ALL_BUSI_PANELS_ARRAY(calc_method, busi_count) = "50% Gross Inc"
    If self_emp_method = "02" Then ALL_BUSI_PANELS_ARRAY(calc_method, busi_count) = "Tax Forms"
    ALL_BUSI_PANELS_ARRAY(mthd_date, busi_count) = replace(method_date, " ", "/")
    If ALL_BUSI_PANELS_ARRAY(mthd_date, busi_count) = "__/__/__" Then ALL_BUSI_PANELS_ARRAY(mthd_date, busi_count) = ""
    ALL_BUSI_PANELS_ARRAY(start_date, busi_count) = replace(income_start, " ", "/")
    ALL_BUSI_PANELS_ARRAY(end_date, busi_count) = replace(income_end, " ", "/")
    If ALL_BUSI_PANELS_ARRAY(end_date, busi_count) = "__/__/__" Then ALL_BUSI_PANELS_ARRAY(end_date, busi_count) = ""

    EMWriteScreen "X", 6, 26
    transmit

    EMReadScreen retro_cash_inc, 8, 9, 43
    EMReadScreen prosp_cash_inc, 8, 9, 59
    EMReadScreen cash_inc_verif, 1, 9, 73
    EMReadScreen retro_cash_exp, 8, 15, 43
    EMReadScreen prosp_cash_exp, 8, 15, 59
    EMReadScreen cash_exp_verif, 1, 15, 73
    ALL_BUSI_PANELS_ARRAY(income_ret_cash, busi_count) = trim(retro_cash_inc)
    If ALL_BUSI_PANELS_ARRAY(income_ret_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_ret_cash, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(income_pro_cash, busi_count) = trim(prosp_cash_inc)
    If ALL_BUSI_PANELS_ARRAY(income_pro_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_pro_cash, busi_count) = "0"
    If cash_inc_verif = "1" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Income Tax Returns"
    If cash_inc_verif = "2" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Receipts of Sales/Purch"
    If cash_inc_verif = "3" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Busi Records/Ledger"
    If cash_inc_verif = "6" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Other Document"
    If cash_inc_verif = "N" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "No Verif Provided"
    If cash_inc_verif = "?" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Delayed Verif"
    If cash_inc_verif = "_" Then ALL_BUSI_PANELS_ARRAY(cash_income_verif, busi_count) = "Blank"
    ALL_BUSI_PANELS_ARRAY(expense_ret_cash, busi_count) = trim(retro_cash_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_ret_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_ret_cash, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(expense_pro_cash, busi_count) = trim(prosp_cash_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_pro_cash, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_pro_cash, busi_count) = "0"
    If cash_exp_verif = "1" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Income Tax Returns"
    If cash_exp_verif = "2" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Receipts of Sales/Purch"
    If cash_exp_verif = "3" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Busi Records/Ledger"
    If cash_exp_verif = "6" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Other Document"
    If cash_exp_verif = "N" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "No Verif Provided"
    If cash_exp_verif = "?" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Delayed Verif"
    If cash_exp_verif = "_" Then ALL_BUSI_PANELS_ARRAY(cash_expense_verif, busi_count) = "Blank"

    EMReadScreen prosp_ive_inc, 8, 10, 59
    EMReadScreen ive_inc_verif, 1, 10, 73
    EMReadScreen prosp_ive_exp, 8, 16, 59
    EMReadScreen ive_exp_verif, 1, 16, 73

    EMReadScreen retro_snap_inc, 8, 11, 43
    EMReadScreen prosp_snap_inc, 8, 11, 59
    EMReadScreen snap_inc_verif, 1, 11, 73
    EMReadScreen retro_snap_exp, 8, 17, 43
    EMReadScreen prosp_snap_exp, 8, 17, 59
    EMReadScreen snap_exp_verif, 1, 17, 73
    ALL_BUSI_PANELS_ARRAY(income_ret_snap, busi_count) = trim(retro_snap_inc)
    If ALL_BUSI_PANELS_ARRAY(income_ret_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_ret_snap, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(income_pro_snap, busi_count) = trim(prosp_snap_inc)
    If ALL_BUSI_PANELS_ARRAY(income_pro_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(income_pro_snap, busi_count) = "0"
    If snap_inc_verif = "1" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Income Tax Returns"
    If snap_inc_verif = "2" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Receipts of Sales/Purch"
    If snap_inc_verif = "3" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Busi Records/Ledger"
    If snap_inc_verif = "4" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Pend Out State Verif"
    If snap_inc_verif = "6" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Other Document"
    If snap_inc_verif = "N" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "No Verif Provided"
    If snap_inc_verif = "?" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Delayed Verif"
    If snap_inc_verif = "_" Then ALL_BUSI_PANELS_ARRAY(snap_income_verif, busi_count) = "Blank"

    ALL_BUSI_PANELS_ARRAY(expense_ret_snap, busi_count) = trim(retro_snap_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_ret_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_ret_snap, busi_count) = "0"
    ALL_BUSI_PANELS_ARRAY(expense_pro_snap, busi_count) = trim(prosp_snap_exp)
    If ALL_BUSI_PANELS_ARRAY(expense_pro_snap, busi_count) = "________" Then ALL_BUSI_PANELS_ARRAY(expense_pro_snap, busi_count) = "0"
    If snap_exp_verif = "1" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Income Tax Returns"
    If snap_exp_verif = "2" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Receipts of Sales/Purch"
    If snap_exp_verif = "3" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Busi Records/Ledger"
    If snap_exp_verif = "4" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Pend Out State Verif"
    If snap_exp_verif = "6" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Other Document"
    If snap_exp_verif = "N" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "No Verif Provided"
    If snap_exp_verif = "?" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Delayed Verif"
    If snap_exp_verif = "_" Then ALL_BUSI_PANELS_ARRAY(snap_expense_verif, busi_count) = "Blank"

    EMReadScreen prosp_hca_inc, 8, 12, 59
    EMReadScreen hca_inc_verif, 1, 12, 73
    EMReadScreen prosp_hca_exp, 8, 18, 59
    EMReadScreen hca_exp_verif, 1, 18, 73

    EMReadScreen prosp_hcb_inc, 8, 13, 59
    EMReadScreen hcb_inc_verif, 1, 13, 73
    EMReadScreen prosp_hcb_exp, 8, 19, 59
    EMReadScreen hcb_exp_verif, 1, 19, 73

    ALL_BUSI_PANELS_ARRAY(budget_explain, busi_count) = ""
    PF3

end function

function read_EATS_panel()
    call navigate_to_MAXIS_screen("STAT", "EATS")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN
        If UBound(ALL_MEMBERS_ARRAY, 2) = 0 Then EATS = "Single member case, EATS panel is not needed,"
        exit function		'Exits out if there's no panel info
    End If
    EMReadScreen all_eat_together, 1, 4, 72
    If all_eat_together = "Y" Then
        EATS = "All clients on this case purchase and prepare food together."
    Else
        EATS = "SNAP unit p/p sep from memb(s):"
        EMReadScreen group_one, 40, 13, 39
        EMReadScreen group_two, 40, 14, 39
        EMReadScreen group_three, 40, 15, 39
        EMReadScreen group_four, 40, 16, 39
        EMReadScreen group_five, 40, 17, 39

        group_one = replace(group_one, "__", "")
        group_two = replace(group_two, "__", "")
        group_three = replace(group_three, "__", "")
        group_four = replace(group_four, "__", "")
        group_five = replace(group_five, "__", "")

        group_one = trim(group_one)
        group_two = trim(group_two)
        group_three = trim(group_three)
        group_four = trim(group_four)
        group_five = trim(group_five)

        If group_one <> "" Then
            EMReadScreen group_one_no, 2, 13, 28
            group_one = replace(group_one, "  ", ", ")
            EATS = EATS & "Eating group " & group_one_no & " with memb(s) " & group_one
        End If
        If group_two <> "" Then
            EMReadScreen group_two_no, 2, 13, 28
            group_two = replace(group_two, "  ", ", ")
            EATS = EATS & "; Eating group " & group_two_no & " with memb(s) " & group_two
        End If
        If group_three <> "" Then
            EMReadScreen group_three_no, 2, 13, 28
            group_three = replace(group_three, "  ", ", ")
            EATS = EATS & "; Eating group " & group_three_no & " with memb(s) " & group_three
        End If
        If group_four <> "" Then
            EMReadScreen group_four_no, 2, 13, 28
            group_four = replace(group_four, "  ", ", ")
            EATS = EATS & "; Eating group " & group_four_no & " with memb(s) " & group_four
        End If
        If group_five <> "" Then
            EMReadScreen group_five_no, 2, 13, 28
            group_five = replace(group_five, "  ", ", ")
            EATS = EATS & "; Eating group " & group_five_no & " with memb(s) " & group_five
        End If

    End If
end function

function read_HEST_panel()
    hest_information = ""
    call navigate_to_MAXIS_screen("STAT", "HEST")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN
        hest_information = "NONE - $0"
    ELSE
        EMReadScreen prosp_heat_air, 1, 13, 60
        EMReadScreen prosp_electric, 1, 14, 60
        EMReadScreen prosp_phone, 1, 15, 60
        combined_electric_and_phone_amt = electric_amt + phone_amt

        If prosp_heat_air = "Y" Then
            hest_information = "AC/Heat - Full $" & heat_AC_amt
        ElseIf prosp_electric = "Y" Then
            If prosp_phone = "Y" Then
                hest_information = "Electric and Phone - $" & combined_electric_and_phone_amt
            Else
                hest_information = "Electric ONLY - $" & electric_amt
            End If
        ElseIf prosp_phone = "Y" Then
            hest_information = "Phone ONLY - $" & phone_amt
        End If
    END IF
end function

function read_JOBS_panel()
    EMReadScreen JOBS_month, 5, 20, 55									'reads Footer month
    JOBS_month = replace(JOBS_month, " ", "/")					'Cleans up the read number by putting a / in place of the blank space between MM YY
    EMReadScreen JOBS_type, 30, 7, 42										'Reads up name of the employer and then cleans it up
    JOBS_type = replace(JOBS_type, "_", ""	)
    JOBS_type = trim(JOBS_type)
    JOBS_type = split(JOBS_type)
    For each JOBS_part in JOBS_type											'Correcting case on the name of the employer as it reads in all CAPS
        If JOBS_part <> "" then
            first_letter = ucase(left(JOBS_part, 1))
            other_letters = LCase(right(JOBS_part, len(JOBS_part) -1))
            new_JOBS_type = new_JOBS_type & first_letter & other_letters & " "
        End if
    Next
    ALL_JOBS_PANELS_ARRAY(employer_name, job_count) = new_JOBS_type
    EMReadScreen jobs_hourly_wage, 6, 6, 75   'reading hourly wage field
    ALL_JOBS_PANELS_ARRAY(hrly_wage, job_count) = replace(jobs_hourly_wage, "_", "")   'trimming any underscores

    ' Navigates to the FS PIC
    EMWriteScreen "X", 19, 38
    transmit
    EMReadScreen SNAP_JOBS_amt, 8, 17, 56
    ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, job_count) = trim(SNAP_JOBS_amt)
    EMReadScreen jobs_SNAP_prospective_amt, 8, 18, 56
    ALL_JOBS_PANELS_ARRAY(pic_prosp_income, job_count) = trim(jobs_SNAP_prospective_amt)  'prospective amount from PIC screen
    EMReadScreen snap_pay_frequency, 1, 5, 64
    ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = snap_pay_frequency
    EMReadScreen date_of_pic_calc, 8, 5, 34
    ALL_JOBS_PANELS_ARRAY(pic_calc_date, job_count) = replace(date_of_pic_calc, " ", "/")
    transmit
    'Navigats to GRH PIC
    EMReadscreen GRH_PIC_check, 3, 19, 73 	'This must check to see if the GRH PIC is there or not. If fun on months 06/16 and before it will cause an error if it pf3s on the home panel.
    IF GRH_PIC_check = "GRH" THEN
    	EMWriteScreen "X", 19, 71
    	transmit
    	EMReadScreen GRH_JOBS_pay_amt, 8, 16, 69
    	GRH_JOBS_pay_amt = trim(GRH_JOBS_pay_amt)
        EMReadScreen GRH_JOBS_total_amt, 8, 17, 69
        GRH_JOBS_total_amt = trim(GRH_JOBS_total_amt)
    	EMReadScreen GRH_pay_frequency, 1, 3, 63
    	EMReadScreen GRH_date_of_pic_calc, 8, 3, 30
    	GRH_date_of_pic_calc = replace(GRH_date_of_pic_calc, " ", "/")
        ALL_JOBS_PANELS_ARRAY(grh_calc_date, job_count) = GRH_date_of_pic_calc
        ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = GRH_pay_frequency
        ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, job_count) = GRH_JOBS_pay_amt
        ALL_JOBS_PANELS_ARRAY(grh_prosp_income, job_count) = GRH_JOBS_total_amt
    	PF3
    END IF
    '  Reads the information on the retro side of JOBS
    EMReadScreen retro_JOBS_amt, 8, 17, 38
    EMReadScreen retro_JOBS_hrs, 3, 18, 43
    retro_JOBS_hrs = replace(retro_JOBS_hrs, "_", "")
    ALL_JOBS_PANELS_ARRAY(job_retro_income, job_count) = trim(retro_JOBS_amt)
    ALL_JOBS_PANELS_ARRAY(retro_hours, job_count) = trim(retro_JOBS_hrs)

    '  Reads the information on the prospective side of JOBS
    EMReadScreen prospective_JOBS_amt, 8, 17, 67
    EMReadScreen prosp_JOBS_hrs, 3, 18, 72
    prosp_JOBS_hrs = replace(prosp_JOBS_hrs, "_", "")
    ALL_JOBS_PANELS_ARRAY(job_prosp_income, job_count) = trim(prospective_JOBS_amt)
    ALL_JOBS_PANELS_ARRAY(prosp_hours, job_count) = trim(prosp_JOBS_hrs)

    '  Reads the information about health care off of HC Income Estimator
    EMReadScreen pay_frequency, 1, 18, 35
    ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = pay_frequency
    EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
    IF HC_income_est_check = "Est" Then 'this is the old position
    	EMWriteScreen "X", 19, 54
    ELSE								'this is the new position
    	EMWriteScreen "X", 19, 48
    END IF
    transmit
    EMReadScreen HC_JOBS_amt, 8, 11, 63
    HC_JOBS_amt = trim(HC_JOBS_amt)
    transmit

    EMReadScreen JOBS_ver, 25, 6, 36
    ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = trim(JOBS_ver)
    If ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = "" Then
        EMReadScreen JOBS_ver, 1, 6, 34
        If JOBS_ver = "?" Then ALL_JOBS_PANELS_ARRAY(verif_code, job_count) = "Delayed"
    End If
    EMReadScreen JOBS_income_end_date, 8, 9, 49
    'This now cleans up the variables converting codes read from the panel into words for the final variable to be used in the output.
    If JOBS_income_end_date <> "__ __ __" then
        JOBS_income_end_date = replace(JOBS_income_end_date, " ", "/")
        ALL_JOBS_PANELS_ARRAY(job_prosp_income, job_count) = "0.00"
        ALL_JOBS_PANELS_ARRAY(prosp_hours, job_count) = "0"
    End If
    If IsDate(JOBS_income_end_date) = True then ALL_JOBS_PANELS_ARRAY(budget_explain, job_count) = "Income ended " & JOBS_income_end_date & ".; "

    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "4" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Weekly"
    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "3" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Biweekly"
    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "2" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Semi-Monthly"
    If ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "1" Then ALL_JOBS_PANELS_ARRAY(main_pay_freq, job_count) = "Monthly"

    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "4" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Weekly"
    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "3" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Biweekly"
    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "2" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Semi-Monthly"
    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "1" Then ALL_JOBS_PANELS_ARRAY(pic_pay_freq, job_count) = "Monthly"

    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "4" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Weekly"
    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "3" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Biweekly"
    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "2" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Semi-Monthly"
    If ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "1" Then ALL_JOBS_PANELS_ARRAY(grh_pay_freq, job_count) = "Monthly"

end function

function read_SANC_panel()

    call  navigate_to_MAXIS_screen("STAT", "SANC")
    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info
    first_sanc_panel  = true

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE Then
            EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            EMReadScreen SANC_total, 1, 2, 78
            If SANC_total <> 0 then
                EMReadScreen memb_sanc_number, 1, 16, 43
                EMReadScreen case_sanc_number, 1, 17, 43
                EMReadScreen case_compliance_date, 8, 17, 72
                EMReadScreen closed_for_7_sanc, 5, 18, 43
                EMReadScreen closed_for_post_7_sanc, 5, 19, 43


                If closed_for_7_sanc = "     " Then
                    If first_sanc_panel = true Then notes_on_sanction = notes_on_sanction & "Total case sanctions: " & case_sanc_number & ".; "

                    notes_on_sanction = notes_on_sanction & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has incurred " & memb_sanc_number & " sanctions.; "
                Else
                    closed_for_7_sanc = replace(closed_for_7_sanc, " ", "/")
                    case_compliance_date = replace(case_compliance_date, " ", "/")
                    If first_sanc_panel = true Then
                        notes_on_sanction = notes_on_sanction & "Total case sanctions: 7. Case was closed for 7th sanction " & closed_for_7_sanc & ".; "
                        ' MsgBox "Case compliance Date - " & case_compliance_date & vbNewLine & "Is Date - " & IsDate(case_compliance_date)
                        If IsDate(case_compliance_date) = True Then notes_on_sanction = notes_on_sanction & "Case came into commpliance after closure for sanction on " & case_compliance_date & ".; "
                        If closed_for_post_7_sanc <> "     " Then
                            closed_for_post_7_sanc = replace(closed_for_post_7_sanc, " ", "/")
                            notes_on_sanction = notes_on_sanction & "Case clossed for 2nd Post-7th saction " & closed_for_post_7_sanc & ".; "
                        End If
                    End If
                    If memb_sanc_number = "6" Then
                        notes_on_sanction = notes_on_sanction & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has incurred 7 sanctions and was closed for 7th sanction " & closed_for_7_sanc & ".; "
                    Else
                        notes_on_sanction = notes_on_sanction & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has incurred " & memb_sanc_number & " sanctions.; "
                    End If


                End If

                first_sanc_panel = false
            End If
        End If
    Next
end function

function read_SHEL_panel()

    call navigate_to_MAXIS_screen("STAT", "SHEL")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen SHEL_total, 1, 2, 78
        SHEL_total = SHEL_total & ""
        If SHEL_total <>"0" then
            member_number_designation = "Member " & HH_member & "- "
            row = 11
            ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE
            Do
                EMReadScreen SHEL_HUD_code, 1, 6, 46
                EMReadScreen SHEL_share_code, 1, 6, 64

                If SHEL_HUD_code = "Y" Then ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "Yes"
                If SHEL_HUD_code = "N" Then ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "No"
                If SHEL_share_code = "Y" Then ALL_MEMBERS_ARRAY(shel_shared, each_member) = "Yes"
                If SHEL_share_code = "N" Then ALL_MEMBERS_ARRAY(shel_shared, each_member) = "No"

                EmReadScreen SHEL_retro_amount, 8, row, 37
                EMReadScreen SHEL_prosp_amount, 8, row, 56
                If SHEL_retro_amount <> "________" OR SHEL_prosp_amount <> "________" then
                    EMReadScreen SHEL_retro_proof, 2, row, 48
                    EMReadScreen SHEL_prosp_proof, 2, row, 67

                    If SHEL_prosp_amount = "________" Then SHEL_prosp_amount = 0
                    SHEL_prosp_amount = trim(SHEL_prosp_amount)
                    SHEL_prosp_amount = SHEL_prosp_amount * 1

                    If SHEL_retro_amount = "________" Then SHEL_retro_amount = 0
                    SHEL_retro_amount = trim(SHEL_retro_amount)
                    SHEL_retro_amount = SHEL_retro_amount * 1

                    If SHEL_retro_proof = "__" Then SHEL_retro_proof = "Blank"
                    If SHEL_prosp_proof = "__" Then SHEL_prosp_proof = "Blank"

                    If SHEL_retro_proof = "SF" Then SHEL_retro_proof = "SF - Shelter Form"
                    If SHEL_prosp_proof = "SF" Then SHEL_prosp_proof = "SF - Shelter Form"
                    If SHEL_retro_proof = "LE" Then SHEL_retro_proof = "LE - Lease"
                    If SHEL_prosp_proof = "LE" Then SHEL_prosp_proof = "LE - Lease"
                    If SHEL_retro_proof = "RE" Then SHEL_retro_proof = "RE - Rent Receipt"
                    If SHEL_prosp_proof = "RE" Then SHEL_prosp_proof = "RE - Rent Receipt"
                    If SHEL_retro_proof = "BI" Then SHEL_retro_proof = "BI - Billing Stmt"
                    If SHEL_prosp_proof = "BI" Then SHEL_prosp_proof = "BI - Billing Stmt"
                    If SHEL_retro_proof = "MO" Then SHEL_retro_proof = "MO - Mort Pmt Book"
                    If SHEL_prosp_proof = "MO" Then SHEL_prosp_proof = "MO - Mort Pmt Book"
                    If SHEL_retro_proof = "CD" Then SHEL_retro_proof = "CD - Ctrct For Deed"
                    If SHEL_prosp_proof = "CD" Then SHEL_prosp_proof = "CD - Ctrct For Deed"
                    If SHEL_retro_proof = "TX" Then SHEL_retro_proof = "TX - Prop Tax Stmt"
                    If SHEL_prosp_proof = "TX" Then SHEL_prosp_proof = "TX - Prop Tax Stmt"
                    If SHEL_retro_proof = "OT" Then SHEL_retro_proof = "OT - Other Doc"
                    If SHEL_prosp_proof = "OT" Then SHEL_prosp_proof = "OT - Other Doc"
                    If SHEL_retro_proof = "NC" Then SHEL_retro_proof = "NC - Change - Neg Impact"
                    If SHEL_prosp_proof = "NC" Then SHEL_prosp_proof = "NC - Change - Neg Impact"
                    If SHEL_retro_proof = "PC" Then SHEL_retro_proof = "PC - Change - Pos Impact"
                    If SHEL_prosp_proof = "PC" Then SHEL_prosp_proof = "PC - Change - Pos Impact"
                    If SHEL_retro_proof = "NO" Then SHEL_retro_proof = "NO - No Verif"
                    If SHEL_prosp_proof = "NO" Then SHEL_prosp_proof = "NO - No Verif"
                    If SHEL_retro_proof = "?_" Then SHEL_retro_proof = "? - Delayed Verif"
                    If SHEL_prosp_proof = "?_" Then SHEL_prosp_proof = "? - Delayed Verif"

                    If row = 11 Then
                        ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 12 Then
                        ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 13 Then
                        ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 14 Then
                        ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 15 Then
                        ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 16 Then
                        ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 17 Then
                        ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) = SHEL_prosp_proof
                    ElseIf row = 18 Then
                        ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) = SHEL_retro_amount
                        ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) = SHEL_retro_proof
                        ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) = SHEL_prosp_amount
                        ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) = SHEL_prosp_proof
                    End If

                    'ADD Reading of panel and saving to the array here'
                End if
                row = row + 1
            Loop until row = 19
        Else
            ALL_MEMBERS_ARRAY(shel_exists, each_member) = False
        End if
        SHEL_expense = ""
    Next
end function

function read_TIME_panel()
    call  navigate_to_MAXIS_screen("STAT", "TIME")
    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE Then
            EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            EMReadScreen TIME_total, 1, 2, 78
            If TIME_total <> 0 then
                EMReadScreen fed_tanf_months, 3, 17, 31
                EMReadScreen state_tanf_months, 3, 17, 53
                EMReadScreen total_tanf_months, 3, 17, 69
                EMReadScreen banked_tanf_months, 3, 19, 16
                EMReadScreen memb_ext_code, 2, 19, 31
                EMReadScreen memb_ext_total, 3, 19, 69

                fed_tanf_months = trim(fed_tanf_months)
                state_tanf_months = trim(state_tanf_months)
                total_tanf_months = trim(total_tanf_months)
                banked_tanf_months = trim(banked_tanf_months)
                memb_ext_total = trim(memb_ext_total)

                used_tanf = total_tanf_months * 1
                tanf_left = 60 - total_tanf_months
                If tanf_left < 0 Then tanf_left = 0

                If memb_ext_code = "01" Then memb_ext_info = "Ill or Incapacitated for more than 30 days"
                If memb_ext_code = "02" Then memb_ext_info = "Care of someone who is Ill or Incapacitated"
                If memb_ext_code = "03" Then memb_ext_info = "Care of someone with Special Medical Criteria"
                If memb_ext_code = "05" Then memb_ext_info = "Unemployable"
                If memb_ext_code = "06" Then memb_ext_info = "Low IQ"
                If memb_ext_code = "07" Then memb_ext_info = "Learning Disabled"
                If memb_ext_code = "08" Then memb_ext_info = "Employed 30+ hours per week (1 caregiver HH)"
                If memb_ext_code = "09" Then memb_ext_info = "Employed 55+ hours per week (2 caregived HH)"
                If memb_ext_code = "10" Then memb_ext_info = "Family Violence"
                If memb_ext_code = "11" Then memb_ext_info = "Developmental Disabilities"
                If memb_ext_code = "12" Then memb_ext_info = "Mental Illness"
                If memb_ext_code = "NO" Then memb_ext_info = "NONE"
                If memb_ext_code = "AP" Then memb_ext_info = "Appeal in Process"
                If memb_ext_code = "__" Then memb_ext_info = ""

                notes_on_time = notes_on_time & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " has used a total of " & total_tanf_months & " TANF months (" & fed_tanf_months & " Federal and " & state_tanf_months & "State) and has " & tanf_left & " TANF months remaining.; "
                If banked_tanf_months <> "0" Then notes_on_time = notes_on_time & "Memb " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " - " & banked_tanf_months & " TANF Banked Months.; "
            End If
        End If
    Next
end function

function read_UNEA_panel()
    call navigate_to_MAXIS_screen("STAT", "UNEA")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info

    If variable_written_to <> "" then variable_written_to = variable_written_to & "; "
    unea_array_counter = 0
    For each HH_member in HH_member_array
        EMWriteScreen HH_member, 20, 76
        EMWriteScreen "01", 20, 79
        transmit
        EMReadScreen UNEA_total, 1, 2, 78

        ReDim Preserve UNEA_INCOME_ARRAY(budget_notes, unea_array_counter)
        UNEA_INCOME_ARRAY(UC_exists, unea_array_counter) = FALSE
        UNEA_INCOME_ARRAY(CS_exists, unea_array_counter) = FALSE
        UNEA_INCOME_ARRAY(SSA_exists, unea_array_counter) = FALSE
        UNEA_INCOME_ARRAY(memb_numb, unea_array_counter) = HH_member
        If UNEA_total <> 0 then
            Do
                EMReadScreen income_type, 2, 5, 37

                EMReadScreen panel_month, 5, 20, 55
                panel_month = replace(panel_month, " ", "/")
                UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) = panel_month

                EMReadScreen UNEA_ver, 1, 5, 65
                If UNEA_ver = "1" Then UNEA_ver = "Copy of Checks"
                If UNEA_ver = "2" Then UNEA_ver = "Award Letters"
                If UNEA_ver = "3" Then UNEA_ver = "System Initiated Verif"
                If UNEA_ver = "4" Then UNEA_ver = "Colateral Statement"
                If UNEA_ver = "5" Then UNEA_ver = "Pend Out State Verif"
                If UNEA_ver = "6" Then UNEA_ver = "Other Document"
                If UNEA_ver = "7" Then UNEA_ver = "Worker Initiated Verif"
                If UNEA_ver = "8" Then UNEA_ver = "RI Stubs"
                If UNEA_ver = "N" Then UNEA_ver = "No Verif"
                If UNEA_ver = "?" Then UNEA_ver = "Delayed"
                If UNEA_ver = "_" Then UNEA_ver = "Blank"

                EMReadScreen UNEA_income_end_date, 8, 7, 68
                If UNEA_income_end_date <> "__ __ __" then UNEA_income_end_date = replace(UNEA_income_end_date, " ", "/")
                EMReadScreen UNEA_income_start_date, 8, 7, 37
                If UNEA_income_start_date <> "__ __ __" then UNEA_income_start_date = replace(UNEA_income_start_date, " ", "/")

                EMReadScreen prosp_amt, 8, 18, 68
                prosp_amt = trim(prosp_amt)
                EMReadScreen retro_amt, 8, 18, 39
                retro_amt = trim(retro_amt)

                EMWriteScreen "X", 10, 26
                transmit
                EMReadScreen SNAP_UNEA_amt, 8, 18, 56
                SNAP_UNEA_amt = trim(SNAP_UNEA_amt)
                EMReadScreen snap_pay_frequency, 1, 5, 64
                EMReadScreen date_of_pic_calc, 8, 5, 34
                date_of_pic_calc = replace(date_of_pic_calc, " ", "/")
                transmit

                If prosp_amt = "" Then prosp_amt = 0
                prosp_amt = prosp_amt * 1
                If retro_amt = "" Then retro_amt = 0
                retro_amt = retro_amt * 1
                If SNAP_UNEA_amt = "" Then SNAP_UNEA_amt = 0
                SNAP_UNEA_amt = SNAP_UNEA_amt * 1

                IF snap_pay_frequency = "1" THEN snap_pay_frequency = "monthly"
                IF snap_pay_frequency = "2" THEN snap_pay_frequency = "semimonthly"
                IF snap_pay_frequency = "3" THEN snap_pay_frequency = "biweekly"
                IF snap_pay_frequency = "4" THEN snap_pay_frequency = "weekly"
                IF snap_pay_frequency = "5" THEN snap_pay_frequency = "non-monthly"

                variable_name_for_UNEA = variable_name_for_UNEA & "UNEA from " & trim(UNEA_type) & ", " & UNEA_month  & " amts:; "
                determined_unea_income = determined_unea_income + SNAP_UNEA_amt

                If SNAP_UNEA_amt <> 0 THEN variable_name_for_UNEA = variable_name_for_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                If retro_UNEA_amt <> 0 THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
                If prosp_UNEA_amt <> 0 THEN variable_name_for_UNEA = variable_name_for_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
                'Leaving out HC income estimator if footer month is not Current month + 1
                If UNEA_ver = "N" or UNEA_ver = "?" then variable_name_for_UNEA = variable_name_for_UNEA & "- No proof provided for this panel; "

                If income_type = "01" or income_type = "02" or income_type = "03" or income_type = "44" Then
                    If IsDate(UNEA_income_end_date) = TRUE Then
                        If income_type = "01" or income_type = "02" Then ssa_type_for_note = "RSDI"
                        If income_type = "03" Then ssa_type_for_note = "SSI"
                        If income_type = "44" then ssa_type_for_note = "Excess Calculation of"
                        notes_on_ssa_income = notes_on_ssa_income & ssa_type_for_note & " income for Memb " & HH_member & " ended on " & UNEA_income_end_date & ". Verification: " & UNEA_ver & "; "
                    Else
                        UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", SSA Income"
                        UNEA_INCOME_ARRAY(SSA_exists, unea_array_counter) = TRUE

                        If income_type = "01" or income_type = "02" Then
                            UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) + prosp_amt
                            'TODO - The NOTES area of the array here does not work well for a case that has a person with more than one RSDI panel. ENHANCEMENT
                            If income_type = "01" Then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "RSDI is Disability Income.; "
                            If SNAP_checkbox = checked and prosp_amt <> SNAP_UNEA_amt Then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "SNAP budgeted Inomce =  $" & SNAP_UNEA_amt & "; "
                           UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        End If
                        If income_type = "03" Then
                            UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) + prosp_amt
                            If SNAP_checkbox = checked and prosp_amt <> SNAP_UNEA_amt Then UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "SNAP budgeted Inomce =  $" & SNAP_UNEA_amt & "; "
                            UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        End If
                    End If
                ElseIf income_type = "14" Then
                    If IsDate(UNEA_income_end_date) = TRUE Then
                        other_uc_income_notes = other_uc_income_notes & "Unemployment Income for Memb " & HH_member & " ended on " & UNEA_income_end_date & ". Verification: " & UNEA_ver & ".; "
                    Else
                        UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", Unemployment"
                        UNEA_INCOME_ARRAY(UC_exists, unea_array_counter) = TRUE

                        UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) + retro_amt
                        UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) + prosp_amt
                        UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) + SNAP_UNEA_amt

                        EMReadScreen pay_day, 8, 13, 68
                        pay_day = trim(pay_day)
                        If pay_day = "" Then pay_day = 0
                        pay_day = pay_day * 1

                        UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) + pay_day
                       UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                        If IsDate(UNEA_income_start_date) = TRUE Then
                            If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                        End If
                        UNEA_INCOME_ARRAY(UNEA_UC_start_date, unea_array_counter) = UNEA_income_start_date
                    End If
                ElseIf income_type = "08" or income_type = "36" or income_type = "39" or income_type = "43" Then
                    If IsDate(UNEA_income_end_date) = TRUE Then
                        If income_type = "08" Then cs_type_for_note = "Direct Child Support"
                        If income_type = "36" Then cs_type_for_note = "Disbursed Child Support"
                        If income_type = "39" Then cs_type_for_note = "Disbursed Child Support Arrears"
                        notes_on_cses = notes_on_cses & cs_type_for_note & " income for Memb " & HH_member & " ended on " & UNEA_income_end_date & ". Verification: " & UNEA_ver & ".; "
                    Else
                        UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", Child Support"
                        UNEA_INCOME_ARRAY(CS_exists, unea_array_counter) = TRUE

                        If income_type = "08" Then
                            UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) + prosp_amt
                            UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        ElseIf income_type = "36" Then
                            UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) + prosp_amt
                            UNEA_INCOME_ARRAY(disb_CS_prosp_budg, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_prosp_budg, unea_array_counter) + SNAP_UNEA_amt
                            UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        ElseIf income_type = "39" Then
                            UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) + prosp_amt
                            UNEA_INCOME_ARRAY(disb_CS_arrears_budg, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_arrears_budg, unea_array_counter) + SNAP_UNEA_amt
                            UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) & "Verif: " & UNEA_ver & "; "
                            If IsDate(UNEA_income_start_date) = TRUE Then
                                If DateDiff("m", UNEA_income_start_date, date) < 6 Then UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                            End If
                            If IsDate(UNEA_income_end_date) = True then UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) = UNEA_INCOME_ARRAY(disb_cs_arrears_notes, unea_array_counter) & "Income ended " & UNEA_income_end_date & "; "
                        End If
                    End If
                ElseIf income_type = "11" or income_type = "12" or income_type = "13" or income_type = "38" Then
                    If income_type = "11" Then income_detail = "Disability Benefit"
                    If income_type = "12" Then income_detail = "Pension"
                    If income_type = "13" Then income_detail = "Other"
                    If income_type = "38" Then income_detail = "Aid & Attendance"

                    UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", VA - " & income_detail

                    notes_on_VA_income = notes_on_VA_income & "; Member " & HH_member & "unearned income from VA (" & income_detail & "), verif: " & UNEA_ver & ", " & UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) & " amts:; "
                    If SNAP_UNEA_amt <> "" THEN notes_on_VA_income = notes_on_VA_income & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                    If retro_UNEA_amt <> "" THEN notes_on_VA_income = notes_on_VA_income & "- Retrospective: $" & retro_UNEA_amt & " total; "
                    If prosp_UNEA_amt <> "" THEN notes_on_VA_income = notes_on_VA_income & "- Prospective: $" & prosp_UNEA_amt & " total; "
                    If IsDate(UNEA_income_start_date) = TRUE Then
                        If DateDiff("m", UNEA_income_start_date, date) < 6 Then notes_on_VA_income = notes_on_VA_income & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                    End If
                    If IsDate(UNEA_income_end_date) = True then notes_on_VA_income = notes_on_VA_income & "Income ended " & UNEA_income_end_date & "; "
                ElseIf income_type = "15" Then
                    UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", Worker's Comp"

                    notes_on_WC_income = notes_on_WC_income & "; Member " & HH_member & "unearned income from Worker's Comp, verif: " & UNEA_ver & ", " & UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) & " amts:; "
                    If SNAP_UNEA_amt <> "" THEN notes_on_WC_income = notes_on_WC_income & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                    If retro_UNEA_amt <> "" THEN notes_on_WC_income = notes_on_WC_income & "- Retrospective: $" & retro_UNEA_amt & " total; "
                    If prosp_UNEA_amt <> "" THEN notes_on_WC_income = notes_on_WC_income & "- Prospective: $" & prosp_UNEA_amt & " total; "
                    If IsDate(UNEA_income_start_date) = TRUE Then
                        If DateDiff("m", UNEA_income_start_date, date) < 6 Then notes_on_WC_income = notes_on_WC_income & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                    End If
                    If IsDate(UNEA_income_end_date) = True then notes_on_WC_income = notes_on_WC_income & "Income ended " & UNEA_income_end_date & "; "

                Else
                    If income_type = "06" Then income_type = "Public Assistance not in MN"
                    If income_type = "19" or income_type = "21" Then income_type = "Foster Care"
                    If income_type = "20" or income_type = "22" Then income_type = "Foster Care (not req FS)"
                    If income_type = "16" Then income_type = "Railroad Retirement"
                    If income_type = "17" Then income_type = "Retirement"
                    If income_type = "35" or income_type = "37" or income_type = "40" Then income_type = "Spousal Support"
                    If income_type = "18" Then income_type = "Military Entitlement"
                    If income_type = "23" Then income_type = "Dividends"
                    If income_type = "24" Then income_type = "Interest"
                    If income_type = "25" Then income_type = "Prizes and Gifts"
                    If income_type = "26" Then income_type = "Strike Benefit"
                    If income_type = "27" Then income_type = "Contract for Deed"
                    If income_type = "28" Then income_type = "Illegal Income"
                    If income_type = "29" Then income_type = "Other Countable"
                    If income_type = "30" Then income_type = "Infreq Irreg"
                    If income_type = "31" Then income_type = "Other FS Only"
                    If income_type = "45" Then income_type = "County 88 Gaming"
                    If income_type = "47" Then income_type = "Tribal Income"
                    If income_type = "48" Then income_type = "Trust Income"
                    If income_type = "49" Then income_type = "Non-Recurring"

                    UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) & ", " & income_type

                    notes_on_other_UNEA = notes_on_other_UNEA & "; Member " & HH_member & "unearned income from " & income_type & ", verif: " & UNEA_ver & ", " & UNEA_INCOME_ARRAY(UNEA_month, unea_array_counter) & " amts:; "
                    If SNAP_UNEA_amt <> "" THEN notes_on_other_UNEA = notes_on_other_UNEA & "- PIC: $" & SNAP_UNEA_amt & "/" & snap_pay_frequency & ", calculated " & date_of_pic_calc & "; "
                    If retro_UNEA_amt <> "" THEN notes_on_other_UNEA = notes_on_other_UNEA & "- Retrospective: $" & retro_UNEA_amt & " total; "
                    If prosp_UNEA_amt <> "" THEN notes_on_other_UNEA = notes_on_other_UNEA & "- Prospective: $" & prosp_UNEA_amt & " total; "
                    If IsDate(UNEA_income_start_date) = TRUE Then
                        If DateDiff("m", UNEA_income_start_date, date) < 6 Then notes_on_other_UNEA = notes_on_other_UNEA & "Income started in the past 6 months on " & UNEA_income_start_date & "; "
                    End If
                    If IsDate(UNEA_income_end_date) = True then notes_on_other_UNEA = notes_on_other_UNEA & "Income ended " & UNEA_income_end_date & "; "

                End If

                EMReadScreen UNEA_panel_current, 1, 2, 73
                If cint(UNEA_panel_current) < cint(UNEA_total) then transmit
            Loop until cint(UNEA_panel_current) = cint(UNEA_total)
        End if
        If left(UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter), 2) = ", " Then UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter) = right(UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter), len(UNEA_INCOME_ARRAY(UNEA_type, unea_array_counter)) - 2)

        UNEA_INCOME_ARRAY(UNEA_prosp_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_prosp_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_retro_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_retro_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_SNAP_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SNAP_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(direct_CS_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) = UNEA_INCOME_ARRAY(disb_CS_arrears_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_RSDI_amt, unea_array_counter) & ""
        UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) = UNEA_INCOME_ARRAY(UNEA_SSI_amt, unea_array_counter) & ""

        unea_array_counter = unea_array_counter + 1
    Next

    ' "01 - RSDI, Disa"
    ' "02 - RSDI, No Disa"
    ' "03 - SSI"
    ' "06 - Non-MN PA"
    ' "11 - VA Disability Benefit"
    ' "12 - VA Pension"
    ' "13 - VA Other"
    ' "38 - VA Aid & Attendance"
    ' "14 - Unemployment Insurance"
    ' "15 - Worker's Comp"
    ' "16 - Railroad Retirement"
    ' "17 - Other Retirement"
    ' "18 - Military Entitlement"
    ' "19 - FC Child Requestiong FS"
    ' "20 - FC Child Not Req FS"
    ' "21 - FC Adult Requesting FS"
    ' "22 - FC Adult Not Req FS"
    ' "23 - Dividends"
    ' "24 - Interest"
    ' "25 - Cnt Gifts or Prizes"
    ' "26 - Strike Benefit"
    ' "27 - Contract for Deed"
    ' "28 - Illegal Income"
    ' "29 - Other Countable"
    ' "30 - Infrequent <30 Not Counted"
    ' "31 - Other FS Only"
    '
    ' "08 - Direct Child Support"
    ' "35 - Direct Spousal Support"
    ' "36 - Disbursed Child Support"
    ' "37 - Disbursed Spousal Sup"
    ' "39 - Disbursed CS Arrears"
    ' "40 - Disbursed Spsl Sup Arrears"
    ' "43 - Disbursed Excess CS"
    '
    ' "44 - MSA - Excess Inc for SSI"
    ' "45 - County 88 Gaming"
    ' "47 - Counted Tribal Income"
    ' "48 - Trust Income (CASH)"
    ' "49 - Non-Recurring Income > $60/ptr (CASH)"
end function

function read_WREG_panel()
    call navigate_to_MAXIS_screen("STAT", "WREG")

    'Now it checks for the total number of panels. If there's 0 Of 0 it'll exit the function for you so as to save oodles of time.
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check = "0 Of 0" THEN exit function		'Exits out if there's no panel info
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            ' MsgBox "Member number is " & ALL_MEMBERS_ARRAY(memb_numb, each_member)
            EMWriteScreen ALL_MEMBERS_ARRAY(memb_numb, each_member), 20, 76
            transmit
            EMReadScreen wreg_total, 1, 2, 78
            IF wreg_total <> "0" THEN
                ALL_MEMBERS_ARRAY(wreg_exists, each_member) = TRUE
                EmWriteScreen "X", 13, 57
                transmit
                bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))
                bene_yr_row = 10
                abawd_counted_months = 0
                second_abawd_period = 0
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
                    abawd_months_array = Split(abawd_info_list, ",")

                    'counting and checking for second set of ABAWD months
                    IF is_counted_month = "Y" or is_counted_month = "N" THEN
                        EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
                        second_abawd_period = second_abawd_period + 1				'adding counted months
                        second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
                        second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
                        ALL_MEMBERS_ARRAY(first_second_set, each_member) = second_counted_months_string
                    END IF

                    'declaring & splitting the second set of abawd months array
                    If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
                    second_months_array = Split(second_set_info_list,",")

                    bene_mo_col = bene_mo_col - 4
                    IF bene_mo_col = 15 THEN
                        bene_yr_row = bene_yr_row - 1
                        bene_mo_col = 63
                    END IF
                    month_count = month_count + 1
                LOOP until month_count = 36
                ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = abawd_counted_months & ""
                If ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = "0" Then ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) = ""
                ALL_MEMBERS_ARRAY(list_abawd_mo, each_member) = abawd_info_list
                ALL_MEMBERS_ARRAY(list_second_set, each_member) = second_set_info_list
                PF3

                EmreadScreen read_WREG_status, 2, 8, 50
                If read_WREG_status = "03" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "03  Unfit for Employment"
                If read_WREG_status = "04" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "04  Responsible for Care of Another"
                If read_WREG_status = "05" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "05  Age 60+"
                If read_WREG_status = "06" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "06  Under Age 16"
                If read_WREG_status = "07" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "07  Age 16-17, live w/ parent"
                If read_WREG_status = "08" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "08  Care of Child <6"
                If read_WREG_status = "09" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "09  Employed 30+ hrs/wk"
                If read_WREG_status = "10" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "10  Matching Grant"
                If read_WREG_status = "11" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "11  Unemployment Insurance"
                If read_WREG_status = "12" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "12  Enrolled in School/Training"
                If read_WREG_status = "13" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "13  CD Program"
                If read_WREG_status = "14" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "14  Receiving MFIP"
                If read_WREG_status = "20" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "20  Pend/Receiving DWP"
                If read_WREG_status = "15" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "15  Age 16-17 not live w/ Parent"
                If read_WREG_status = "16" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "16  50-59 Years Old"
                If read_WREG_status = "21" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "21  Care child < 18"
                If read_WREG_status = "17" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "17  Receiving RCA or GA"
                If read_WREG_status = "30" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "30  FSET Participant"
                If read_WREG_status = "02" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "02  Fail FSET Coop"
                If read_WREG_status = "33" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "33  Non-coop being referred"
                If read_WREG_status = "__" THEN  ALL_MEMBERS_ARRAY(clt_wreg_status, each_member) = "__  Blank"

                EmreadScreen read_abawd_status, 2, 13, 50
                If read_abawd_status = "01" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "01  WREG Exempt"
                If read_abawd_status = "02" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "02  Under Age 18"
                If read_abawd_status = "03" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "03  Age 50+"
                If read_abawd_status = "04" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "04  Caregiver of Minor Child"
                If read_abawd_status = "05" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "05  Pregnant"
                If read_abawd_status = "06" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "06  Employed 20+ hrs/wk"
                If read_abawd_status = "07" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "07  Work Experience"
                If read_abawd_status = "08" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "08  Other E and T"
                If read_abawd_status = "09" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "09  Waivered Area"
                IF read_abawd_status = "10" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "10  ABAWD Counted"
                If read_abawd_status = "11" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "11  Second Set"
                If read_abawd_status = "12" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "12  RCA or GA Participant"
                If read_abawd_status = "13" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "13  ABAWD Banked Months"
                If read_abawd_status = "__" THEN  ALL_MEMBERS_ARRAY(clt_abawd_status, each_member) = "__  Blank"

                EMReadScreen read_counter, 1, 14, 50
                If read_counter = "_" Then read_counter = 0
                ALL_MEMBERS_ARRAY(numb_banked_mo, each_member) = read_counter
            End If
        END IF
    Next
End function

function split_string_into_parts(full_string, partial_one, partial_two, partial_three, length_one, length_two)
    If left(full_string, 1) = "*" Then full_string = right(full_string, len(full_string) - 1)
    full_string = trim(full_string)
    If right(full_string, 1) = ";" Then full_string = left(full_string, len(full_string) - 1)
    full_string = trim(full_string)

    If len(full_string) =< length_one Then
        partial_one = full_string
        exit function
    End If

    full_string = replace(full_string, "  ", " ")
    word_array = split(full_string, " ")
    level = 1

    For each word in word_array
        If level = 1 Then
            If len(partial_one & " " & word) > length_one Then level = 2
        ElseIf level = 2 Then
            If partial_three <> "NONE" Then
                If len(partial_two & " " & word) > length_two Then level = 3
            End If
        End If

        If level = 1 Then
            partial_one = partial_one & " " & word
        ElseIf level = 2 Then
            partial_two = partial_two & " " & word
        ElseIf level = 3 Then
            partial_three = partial_three & " " & word
        End If
        '
        ' If len(partial_one & " " & word) < length_one Then
        '     partial_one = partial_one & " " & word
        ' ElseIf partial_three = "NONE" Then
        '     partial_two = partial_two & " " & word
        ' Else
        '     If len(partial_two & " " & word) < length_two Then
        '         partial_two = partial_two & " " & word
        '     Else
        '
        '         partial_three = partial_three & " " & word
        '     End If
        ' End If
        ' MsgBox "Word - " & word & vbNewLine & partial_one & vbNewLine & partial_two & vbNewLine & partial_three
    Next
end function

function update_shel_notes()
    total_shelter_amount = 0
    full_shelter_details = ""
    shelter_details = ""
    shelter_details_two = ""
    shelter_details_three = ""

    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(gather_detail, each_member) = TRUE Then

            If ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE Then
                full_shelter_details = full_shelter_details & "* M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " shelter expense(s): "
                If ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Rent $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Rent retro - $" & ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ".; " & ". "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Rent (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_rent_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Rent (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Lot Rent $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Lot Rent retro - $" & ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Lot Rent (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_lot_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Lot Rent (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Mortgage $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Mortgage retro - $" & ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Mortgage (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Mortgage (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Home Insurance $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Home Insurance retro - $" & ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Home Insurance (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_ins_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Home Insurance (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Property Tax $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Property Tax retro - $" & ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Property Tax (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_tax_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Property Tax (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Room $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Room retro - $" & ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Room (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_room_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Room (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_room_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_room_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Garage $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Garage retro - $" & ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ".; "
                        End If
                        total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Garage (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_garage_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Garage (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, each_member)) - 4) & ".; "
                    total_shelter_amount = total_shelter_amount + ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, each_member)
                End If

                If ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member) <> "Select one" Then
                    If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Select one" Then
                        If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) Then
                            full_shelter_details = full_shelter_details & "Subsidy $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ".; "
                        Else
                            full_shelter_details = full_shelter_details & "change in Subsidy retro - $" & ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member)) - 4) & " prosp - $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ".; "
                        End If
                        'total_shelter_amount = total_shelter_amount - ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)
                    Else
                        full_shelter_details = full_shelter_details & "Subsidy (retro only) $" & ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, each_member)) - 4) & ".; "
                    End If
                ElseIf ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Blank" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "" AND ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member) <> "Select one" Then
                    full_shelter_details = full_shelter_details & "Subsidy (prosp) $" & ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member) & " verif: " & right(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member), len(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, each_member)) - 4) & ".; "
                    'total_shelter_amount = total_shelter_amount - ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, each_member)
                End If
                If ALL_MEMBERS_ARRAY(shel_shared, each_member) = "Yes" Then full_shelter_details = full_shelter_details & "Shelter expense is SHARED. "
                If ALL_MEMBERS_ARRAY(shel_subsudized, each_member) = "Yes" Then full_shelter_details = full_shelter_details & "Shelter expense is subsidized. "
            End If
        End If
    Next

    total_shelter_amount = FormatCurrency(total_shelter_amount)

    ' MsgBOx "Length of full_shelter_details is " & len(full_shelter_details)
    Call split_string_into_parts(full_shelter_details, shelter_details, shelter_details_two, shelter_details_three, 85, 85)
    ' if left(full_shelter_details, 2) = "; " Then full_shelter_details = right(full_shelter_details, len(full_shelter_details) - 2)
    ' If len(full_shelter_details) > 85 Then
    '     shelter_details = left(full_shelter_details, 85)
    '     shelter_details_two = right(full_shelter_details, len(full_shelter_details) - 85)
    ' Else
    '     shelter_details = full_shelter_details
    ' End If
end function

function update_wreg_and_abawd_notes()
    notes_on_wreg = ""
    full_abawd_info = ""
    notes_on_abawd = ""
    notes_on_abawd_two = ""
    notes_on_abawd_three = ""
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        ' MsgBox "Each member - " & each_member & vbNewLine & "M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & vbNewLine & "WREG info - " & ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            If trim(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)) <> "" Then
                notes_on_wreg = notes_on_wreg & "M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & ": WREG - " & right(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member), len(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)) - 4) & " ABAWD - " & right(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), len(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member)) - 4) & "; "
                clt_currently_is = ""
                full_abawd_info = full_abawd_info & "M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
                If left(ALL_MEMBERS_ARRAY(clt_wreg_status, each_member), 2) = "30" Then
                    If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "10" Then clt_currently_is = "ABAWD"
                    If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "11" Then clt_currently_is = "SECOND SET"
                    'If left(ALL_MEMBERS_ARRAY(clt_abawd_status, each_member), 2) = "13" Then clt_currently_is = "BANKED"
                End If
                If clt_currently_is <> "" Then
                    full_abawd_info = full_abawd_info & " currently using " & clt_currently_is & " months."
                End If
                If ALL_MEMBERS_ARRAY(pwe_checkbox, each_member) = checked Then full_abawd_info = full_abawd_info & "; M" & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " is the  SNAP PWE"
                If ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) <> "" OR trim(ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; ABAWD months used: " & ALL_MEMBERS_ARRAY(numb_abawd_used, each_member) & " - " & ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)
                If trim(ALL_MEMBERS_ARRAY(first_second_set, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; 2nd Set used starting: " & ALL_MEMBERS_ARRAY(first_second_set, each_member)
                If trim(ALL_MEMBERS_ARRAY(explain_no_second, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; 2nd Set not available due to: " & ALL_MEMBERS_ARRAY(explain_no_second, each_member)
                'If trim(ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; Banked months used: " & ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)
                If trim(ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)) <> "" Then full_abawd_info = full_abawd_info & "; Notes: " & ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)
                full_abawd_info = full_abawd_info & "; "
            End If
        End If
    Next
    if right(notes_on_wreg, 2) = "; " Then notes_on_wreg = left(notes_on_wreg, len(notes_on_wreg) - 2)

    Call split_string_into_parts(full_abawd_info, notes_on_abawd, notes_on_abawd_two, notes_on_abawd_three, 135, 135)
    ' if right(full_abawd_info, 2) = "; " Then full_abawd_info = left(full_abawd_info, len(full_abawd_info) - 2)
    ' If len(full_abawd_info) > 400 Then
    '     notes_on_abawd = left(full_abawd_info, 400)
    '     notes_on_abawd_two = right(full_abawd_info, len(full_abawd_info) - 400)
    ' Else
    '     notes_on_abawd = full_abawd_info
    ' End If
end function

function verification_dialog()
    If ButtonPressed = verif_button Then
        If second_call <> TRUE Then
            income_source_list = "Select or Type Source"

            For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                If ALL_JOBS_PANELS_ARRAY(employer_name, each_job) <> "" Then income_source_list = income_source_list+chr(9)+"JOB - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
            Next
            For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                If ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) <> "" Then
                    If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then
                        income_source_list = income_source_list+chr(9)+"Self Emp - " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi)
                    Else
                        income_source_list = income_source_list+chr(9)+"Self Employment"
                    End If
                End If
            Next
            employment_source_list = income_source_list
            income_source_list = income_source_list+chr(9)+"Child Support"+chr(9)+"Social Security Income"+chr(9)+"Unemployment Income"+chr(9)+"VA Income"+chr(9)+"Pension"
            income_verif_time = "[Enter Time Frame]"
            bank_verif_time = "[Enter Time Frame]"
            second_call = TRUE
        End If

        Do
            verif_err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 610, 395, "Select Verifications"
              Text 280, 10, 120, 10, "Date Verification Request Form Sent:"
              EditBox 400, 5, 50, 15, verif_req_form_sent_date

              GroupBox 530, 35, 75, 145, "PROGRAM(S):"
              Text 535, 48, 65, 40, "Check all programs that require any of the listed verifications:"
              CheckBox 540, 85, 45, 10, "SNAP", verif_snap_checkbox
              CheckBox 540, 95, 45, 10, "CASH", verif_cash_checkbox
              CheckBox 540, 105, 45, 10, "MFIP", verif_mfip_checkbox
              CheckBox 540, 115, 45, 10, "DWP", verif_dwp_checkbox
              CheckBox 540, 125, 45, 10, "MSA", verif_msa_checkbox
              CheckBox 540, 135, 45, 10, "GA", verif_ga_checkbox
              CheckBox 540, 145, 45, 10, "GRH", verif_grh_checkbox
              CheckBox 540, 155, 45, 10, "EMER", verif_emer_checkbox
              CheckBox 540, 165, 45, 10, "HC", verif_hc_checkbox
              ' CheckBox 540, 50, 45, 10, "SNAP", verif_snap_checkbox

              Groupbox 5, 35, 520, 130, "Personal and Household Information"

              CheckBox 10, 50, 75, 10, "Verification of ID for ", id_verif_checkbox
              ComboBox 90, 45, 150, 45, verification_memb_list, id_verif_memb
              CheckBox 300, 50, 100, 10, "Social Security Number for ", ssn_checkbox
              ComboBox 405, 45, 110, 45, verification_memb_list, ssn_verif_memb

              CheckBox 10, 70, 70, 10, "US Citizenship for ", us_cit_status_checkbox
              ComboBox 85, 65, 150, 45, verification_memb_list, us_cit_verif_memb
              CheckBox 300, 70, 85, 10, "Immigration Status for", imig_status_checkbox
              ComboBox 390, 65, 125, 45, verification_memb_list, imig_verif_memb

              CheckBox 10, 90, 90, 10, "Proof of relationship for ", relationship_checkbox
              ComboBox 105, 85, 150, 45, verification_memb_list, relationship_one_verif_memb
              Text 260, 90, 90, 10, "and"
              ComboBox 280, 85, 150, 45, verification_memb_list, relationship_two_verif_memb

              CheckBox 10, 110, 85, 10, "Student Information for ", student_info_checkbox
              ComboBox 100, 105, 150, 45, verification_memb_list, student_verif_memb
              Text 255, 110, 10, 10, "at"
              EditBox 270, 105, 150, 15, student_verif_source

              CheckBox 10, 130, 85, 10, "Proof of Pregnancy for", preg_checkbox
              ComboBox 100, 125, 150, 45, verification_memb_list, preg_verif_memb

              CheckBox 10, 150, 115, 10, "Illness/Incapacity/Disability for", illness_disability_checkbox
              ComboBox 130, 145, 150, 45, verification_memb_list, disa_verif_memb
              Text 285, 150, 30, 10, "verifying:"
              EditBox 320, 145, 150, 15, disa_verif_type

              GroupBox 5, 165, 520, 50, "Income Information"

              CheckBox 10, 180, 45, 10, "Income for ", income_checkbox
              ComboBox 60, 175, 140, 45, verification_memb_list, income_verif_memb
              Text 205, 180, 15, 10, "from"
              ComboBox 225, 175, 125, 45, income_source_list, income_verif_source
              Text 355, 180, 10, 10, "for"
              EditBox 370, 175, 145, 15, income_verif_time

              CheckBox 10, 200, 85, 10, "Employment Status for ", employment_status_checkbox
              ComboBox 100, 195, 150, 45, verification_memb_list, emp_status_verif_memb
              Text 255, 200, 10, 10, "at"
              ComboBox 270, 195, 150, 45, employment_source_list, emp_status_verif_source

              GroupBox 5, 215, 520, 50, "Expense Information"

              CheckBox 10, 230, 105, 10, "Educational Funds/Costs for", educational_funds_cost_checkbox
              ComboBox 120, 225, 150, 45, verification_memb_list, stin_verif_memb

              CheckBox 10, 250, 65, 10, "Shelter Costs for ", shelter_checkbox
              ComboBox 80, 245, 150, 45, verification_memb_list, shelter_verif_memb
              checkBox 240, 250, 175, 10, "Check here if this verif is NOT MANDATORY", shelter_not_mandatory_checkbox

              GroupBox 5, 265, 600, 30, "Asset Information"

              CheckBox 10, 280, 70, 10, "Bank Account for", bank_account_checkbox
              ComboBox 80, 275, 150, 45, verification_memb_list, bank_verif_memb
              Text 235, 280, 45, 10, "account type"
              ComboBox 285, 275, 145, 45, "Select or Type"+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Certificate of Deposit (CD)"+chr(9)+"Stock"+chr(9)+"Money Market", bank_verif_type
              Text 435, 280, 10, 10, "for"
              EditBox 450, 275, 150, 15, bank_verif_time

              Text 5, 305, 20, 10, "Other:"
              EditBox 30, 300, 570, 15, other_verifs
              Checkbox 10, 320, 200, 10, "Check here to have verifs numbered in the CASE/NOTE.", number_verifs_checkbox
              Checkbox 220, 320, 200, 10, "Check here if there are verifs that have been postponed.", verifs_postponed_checkbox

              ButtonGroup ButtonPressed
                PushButton 485, 10, 50, 15, "FILL", fill_button
                PushButton 540, 10, 60, 15, "Return to Dialog", return_to_dialog_button
              Text 10, 340, 580, 50, verifs_needed
              Text 10, 10, 235, 10, "Check the boxes for any verification you want to add to the CASE/NOTE."
              Text 10, 20, 470, 10, "Note: After you press 'Fill' or 'Return to Dialog' the information from the boxes will fill in the Verification Field and the boxes will be 'unchecked'."
            EndDialog

            dialog Dialog1

            If ButtonPressed = 0 Then
                id_verif_checkbox = unchecked
                us_cit_status_checkbox = unchecked
                imig_status_checkbox = unchecked
                ssn_checkbox = unchecked
                relationship_checkbox = unchecked
                income_checkbox = unchecked
                employment_status_checkbox = unchecked
                student_info_checkbox = unchecked
                educational_funds_cost_checkbox = unchecked
                shelter_checkbox = unchecked
                bank_account_checkbox = unchecked
                preg_checkbox = unchecked
                illness_disability_checkbox = unchecked
            End If
            If ButtonPressed = -1 Then ButtonPressed = fill_button

            If id_verif_checkbox = checked AND (id_verif_memb = "Select or Type Member" OR trim(id_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs ID verified."
            If us_cit_status_checkbox = checked AND (us_cit_verif_memb = "Select or Type Member" OR trim(us_cit_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs citizenship verified."
            If imig_status_checkbox = checked AND (imig_verif_memb = "Select or Type Member" OR trim(imig_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs immigration status verified."
            If ssn_checkbox = checked AND (ssn_verif_memb = "Select or Type Member" OR trim(ssn_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member for which we need social security number."
            If relationship_checkbox = checked Then
                If relationship_one_verif_memb = "Select or Type Member" OR trim(relationship_one_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the two household members whose relationship needs to be verified."
                If relationship_two_verif_memb = "Select or Type Member" OR trim(relationship_two_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the two household members whose relationship needs to be verified."
            End If
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

            If verif_err_msg = "" Then
                If id_verif_checkbox = checked Then
                    If IsNumeric(left(id_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Identity for Memb " & id_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Identity for " & id_verif_memb & ".; "
                    End If
                    id_verif_checkbox = unchecked
                    id_verif_memb = ""
                End If
                If us_cit_status_checkbox = checked Then
                    If IsNumeric(left(us_cit_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "US Citizenship for Memb " & us_cit_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "US Citizenship for " & us_cit_verif_memb & ".; "
                    End If
                    us_cit_status_checkbox = unchecked
                    us_cit_verif_memb = ""
                End If
                If imig_status_checkbox = checked Then
                    If IsNumeric(left(imig_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Immigration documentation for Memb " & imig_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Immigration documentation for " & imig_verif_memb & ".; "
                    End If
                    imig_status_checkbox = unchecked
                    imig_verif_memb = ""
                End If
                If ssn_checkbox = checked Then
                    If IsNumeric(left(ssn_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Social Security number for Memb " & ssn_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Social Security number for " & ssn_verif_memb & ".; "
                    End If
                    ssn_checkbox = unchecked
                    ssn_verif_memb = ""
                End If
                If relationship_checkbox = checked Then
                    If IsNumeric(left(relationship_one_verif_memb, 2)) = TRUE AND IsNumeric(left(relationship_two_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Relationship between Memb " & relationship_one_verif_memb & " and Memb " & relationship_two_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Relationship between " & relationship_one_verif_memb & " and " & relationship_two_verif_memb & ".; "
                    End If
                    relationship_checkbox = unchecked
                    relationship_one_verif_memb = ""
                    relationship_two_verif_memb = ""
                End If
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
            End If

            If ButtonPressed = fill_button Then verif_err_msg = "LOOP" & verif_err_msg
        Loop until verif_err_msg = ""
        ButtonPressed = verif_button
    End If

end function

function run_expedited_determination_script_functionality(xfs_screening, caf_one_income, caf_one_assets, caf_one_rent, caf_one_utilities, determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, approval_date, date_of_application, interview_date, applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, snap_denial_date, snap_denial_explain, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, postponed_verifs_yn, list_postponed_verifs, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, action_due_to_out_of_state_benefits, case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn, delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
    'Add assigning of information here

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
    return_to_main_btn                          = 170

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
    ebt_card_info_btn 	= 1900

    show_pg_amounts = 1
    show_pg_determination = 2
    show_pg_review = 3

    page_display = show_pg_amounts
    Do
    	Do
    		err_msg = ""
    		If page_display = show_pg_determination Then Call determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)
    		If page_display = show_pg_review Then Call determine_actions(case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, is_elig_XFS, snap_denial_date, approval_date, date_of_application, do_we_have_applicant_id, action_due_to_out_of_state_benefits, mn_elig_begin_date, other_snap_state, case_has_previously_postponed_verifs_that_prevent_exp_snap, delay_action_due_to_faci, deny_snap_due_to_faci)

            If determined_income = "" Then determined_income = 0
            If determined_assets = "" Then determined_assets = 0
            If determined_shel = "" Then determined_shel = 0
    		If determined_utilities = "" Then determined_utilities = 0
            If calculated_resources = "" Then calculated_resources = 0
            If calculated_expenses = "" Then calculated_expenses = 0
            If IsNumeric(determined_income) = False Then determined_income = 0
            If IsNumeric(determined_assets) = False Then determined_assets = 0
            If IsNumeric(determined_shel) = False Then determined_shel = 0
            If IsNumeric(determined_utilities) = False Then determined_utilities = 0
            If IsNumeric(calculated_resources) = False Then calculated_resources = 0
            If IsNumeric(calculated_expenses) = False Then calculated_expenses = 0

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
    				GroupBox 5, 105, 390, 125, "Information from STAT"
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
    			    Text 175, 200, 80, 10, "Specifc case situations:"
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

    				' If IsDate(snap_denial_date) = True Then
    				' 	Text 15, 20, 280, 20, "DENIAL has been determined - Case does not meet 'All Other Eligibility Criteria' and Expedited Determination is not needed"
    				'
    				' 	Text 25, 55, 435, 10, "Update MAXIS STAT panels correctly to general results to Deny the Application"
    				' 	Text 25, 70, 435, 10, "Complete the DENIAL and enter a full, detailed CASE/NOTE of the Denial Action and Reasons."
    				' 	Text 25, 85, 435, 10, "Complete ALL PROCESSING before moving on to your next tast. Contact Knowledge Now if you are unsure of a Denial."
    				' ElseIf is_elig_XFS = True Then
    			    ' 	If IsDate(approval_date) = True Then
    				' 		Text 15, 20, 205, 10, "Case appears EXPEDITED and there are NO Delay reasons"
    				'
    				' 	    Text 25, 55, 435, 10, "Update MAXIS STAT panels to generate EXPEDITED SNAP Eligibility Results"
    				' 	    Text 25, 70, 435, 10, "Expedited Package includes " & expedited_package
    				' 	    Text 25, 85, 435, 10, "Approve SNAP Expedited package before moving on to the next task"
    				' 	Else
    				' 	End If
    				' ElseIf is_elig_XFS = False Then
    				' End If
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
    			    PushButton 270, 257, 195, 13, "Information about EBT Cards", ebt_card_info_btn

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
    		    ' CancelButton 500, 365, 50, 15
                PushButton 500, 365, 50, 15, "Return", return_to_main_btn
    		    ' OkButton 500, 350, 50, 15
    		EndDialog

    		Dialog Dialog1
    		cancel_confirmation
    		' MsgBox "1 - ButtonPressed is " & ButtonPressed

    		If ButtonPressed = -1 Then
    			If page_display <> show_pg_review then ButtonPressed = next_btn
    			If page_display = show_pg_review then ButtonPressed = finish_btn
    		End If

    		' If ButtonPressed = income_calc_btn Then Call app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, JOBS_ARRAY, BUSI_ARRAY, UNEA_ARRAY)
    		' If ButtonPressed = asset_calc_btn Then Call app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, ACCOUNTS_ARRAY)
    		' If ButtonPressed = housing_calc_btn Then Call app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)
    		' If ButtonPressed = utility_calc_btn Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
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
    			If postponed_verifs_yn = "Yes" AND trim(list_postponed_verifs) = "" Then err_msg = err_msg * vbCr & "* Since you have Postponed Verifications indicated, list what they are for the NOTE."
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
            If ButtonPressed = return_to_main_btn OR ButtonPressed = income_calc_btn OR ButtonPressed = asset_calc_btn OR ButtonPressed = housing_calc_btn OR ButtonPressed = utility_calc_btn Then
                full_determination_done = False
                err_msg = ""
                show_eight = False
                If ButtonPressed = return_to_main_btn Then show_eight = True
                If ButtonPressed = income_calc_btn Then show_two = True
                If ButtonPressed = asset_calc_btn Then show_seven = True
                If ButtonPressed = housing_calc_btn Then show_six = True
                If ButtonPressed = utility_calc_btn Then show_six = True

            End If

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
			    If ButtonPressed = ebt_card_info_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Accounting.aspx#%E2%80%8B%E2%80%8B%E2%80%8B%E2%80%8B%E2%80%8B%E2%80%8Bprocesses-for-receiving-ebt-cards-at-the-county-offices"

    			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & resource_URL
    		End If



    	Loop until err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
end function

'===========================================================================================================================

'DECLARATIONS ==============================================================================================================
'Constants
'JOBS Array and BUSI Array constants
const memb_numb             = 0
const panel_instance        = 1
const employer_name         = 2
const busi_type             = 2         'for BUSI Array
const estimate_only         = 3
const verif_explain         = 4
const verif_code            = 5
const calc_method           = 5         'for BUSI Array
const info_month            = 6
const hrly_wage             = 7
const mthd_date             = 7          'for BUSI Array'
const main_pay_freq         = 8
const rept_retro_hrs        = 8          'for BUSI Array'
const job_retro_income      = 9
const rept_prosp_hrs        = 9          'for BUSI Array'
const job_prosp_income      = 10
const min_wg_retro_hrs      = 10         'for BUSI Array'
const retro_hours           = 11
const min_wg_prosp_hrs      = 11         'for BUSI Array'
const prosp_hours           = 12
const income_ret_cash       = 12         'for BUSI Array'
const pic_pay_date_income   = 13
const income_pro_cash       = 13         'for BUSI Array'
const pic_pay_freq          = 14
const cash_income_verif     = 14         'for BUSI Array'
const pic_prosp_income      = 15
const expense_ret_cash      = 15         'for BUSI Array'
const pic_calc_date         = 16
const expense_pro_cash      = 16         'for BUSI Array'
const EI_case_note          = 17
const cash_expense_verif    = 17         'for BUSI Array'
const grh_calc_date         = 18
const income_ret_snap       = 18         'for BUSI Array'
const grh_pay_freq          = 19
const income_pro_snap       = 19         'for BUSI Array'
const grh_pay_day_income    = 20
const snap_income_verif     = 20         'for BUSI Array'
const grh_prosp_income      = 21
const expense_ret_snap      = 21         'for BUSI Array'
const expense_pro_snap      = 22         'for BUSI Array'
const snap_expense_verif    = 23         'for BUSI Array'
const method_convo_checkbox = 24         'for BUSI Array'
const start_date            = 25
const end_date              = 26
const busi_desc             = 27         'for BUSI Array'
const busi_structure        = 28         'for BUSI Array'
const share_num             = 29         'for BUSI Array'
const share_denom           = 30         'for BUSI Array'
const partners_in_HH        = 31         'for BUSI Array'
const exp_not_allwd         = 32         'for BUSI Array'
const verif_checkbox        = 33
const verif_added           = 34
const budget_explain        = 35

'Member array constants
const clt_name                  = 1
const clt_age                   = 2
const full_clt                  = 3
const clt_id_verif              = 4
const include_cash_checkbox     = 5
const include_snap_checkbox     = 6
const include_emer_checkbox     = 7
const count_cash_checkbox       = 8
const count_snap_checkbox       = 9
const count_emer_checkbox       = 10
const clt_wreg_status           = 11
const clt_abawd_status          = 12
const pwe_checkbox              = 13
const numb_abawd_used           = 14
const list_abawd_mo             = 15
const first_second_set          = 16
const list_second_set           = 17
const explain_no_second         = 18
const numb_banked_mo            = 19
const clt_abawd_notes           = 20
const shel_exists               = 21
const shel_subsudized           = 22
const shel_shared               = 23
const shel_retro_rent_amt       = 24
const shel_retro_rent_verif     = 25
const shel_prosp_rent_amt       = 26
const shel_prosp_rent_verif     = 27
const shel_retro_lot_amt        = 28
const shel_retro_lot_verif      = 29
const shel_prosp_lot_amt        = 30
const shel_prosp_lot_verif      = 31
const shel_retro_mortgage_amt   = 32
const shel_retro_mortgage_verif = 33
const shel_prosp_mortgage_amt   = 34
const shel_prosp_mortgage_verif = 35
const shel_retro_ins_amt        = 36
const shel_retro_ins_verif      = 37
const shel_prosp_ins_amt        = 38
const shel_prosp_ins_verif      = 39
const shel_retro_tax_amt        = 40
const shel_retro_tax_verif      = 41
const shel_prosp_tax_amt        = 42
const shel_prosp_tax_verif      = 43
const shel_retro_room_amt       = 44
const shel_retro_room_verif     = 45
const shel_prosp_room_amt       = 46
const shel_prosp_room_verif     = 47
const shel_retro_garage_amt     = 48
const shel_retro_garage_verif   = 49
const shel_prosp_garage_amt     = 50
const shel_prosp_garage_verif   = 51
const shel_retro_subsidy_amt    = 52
const shel_retro_subsidy_verif  = 53
const shel_prosp_subsidy_amt    = 54
const shel_prosp_subsidy_verif  = 55
const wreg_exists               = 56
const shel_verif_checkbox       = 57
const shel_verif_added          = 58
const gather_detail             = 59
const id_detail                 = 60
const id_required               = 61
const clt_notes                 = 62

'FOR CS Array'
const UNEA_type                 = 2
const UNEA_month                = 3
const UNEA_verif                = 4
const UNEA_prosp_amt            = 5
const UNEA_retro_amt            = 6
const UNEA_SNAP_amt             = 7
const UNEA_pay_freq             = 8
const UNEA_pic_date_calc        = 9

const UNEA_UC_start_date        = 10
const UNEA_UC_weekly_gross      = 11
const UNEA_UC_counted_ded       = 12
const UNEA_UC_exclude_ded       = 13
const UNEA_UC_weekly_net        = 14
const UNEA_UC_monthly_snap      = 15
const UNEA_UC_retro_amt         = 16
const UNEA_UC_prosp_amt         = 17
const UNEA_UC_notes             = 18
const UNEA_UC_tikl_date         = 19
const UNEA_UC_account_balance   = 20

const direct_CS_amt             = 21
const disb_CS_amt               = 22
const disb_CS_arrears_amt       = 23
const direct_CS_notes           = 24
const disb_CS_notes             = 25
const disb_CS_arrears_notes     = 26
const disb_CS_months            = 27
const disb_CS_prosp_budg        = 28
const disb_CS_arrears_months    = 29
const disb_CS_arrears_budg      = 30

const UNEA_RSDI_amt             = 31
const UNEA_RSDI_notes           = 32
const UNEA_SSI_amt              = 33
const UNEA_SSI_notes            = 34

const UC_exists                 = 35
const CS_exists                 = 36
const SSA_exists                = 37
const calc_button               = 38

const budget_notes              = 39

'Arrays
Dim HH_member_array()
ReDim HH_member_array(0)

Dim ALL_JOBS_PANELS_ARRAY()
ReDim ALL_JOBS_PANELS_ARRAY(budget_explain, 0)

Dim ALL_BUSI_PANELS_ARRAY()
ReDim ALL_BUSI_PANELS_ARRAY(budget_explain, 0)

Dim ALL_MEMBERS_ARRAY()
ReDim ALL_MEMBERS_ARRAY(clt_notes, 0)

Dim UNEA_INCOME_ARRAY()
ReDim UNEA_INCOME_ARRAY(budget_notes, 0)
manual_amount_used = FALSE

'variables
Dim row, col, number_verifs_checkbox, verifs_postponed_checkbox, notes_on_cses
Dim MAXIS_footer_month, MAXIS_footer_year, CASH_checkbox, GRH_checkbox, SNAP_checkbox, EMER_checkbox, HC_checkbox, CAF_form
Dim adult_cash, family_cash, the_process_for_cash, type_of_cash, cash_recert_mo, cash_recert_yr, the_process_for_snap, snap_recert_mo, snap_recert_yr
Dim the_process_for_grh, grh_recert_mo, grh_recert_yr, the_process_for_emer, type_of_emer, multiple_CAF_dates, multiple_interview_dates
Dim the_process_for_hc, hc_recert_mo, hc_recert_yr, application_processing, recert_processing, CAF_datestamp, interview_date, case_details_and_notes_about_process, SNAP_recert_is_likely_24_months, exp_screening_note_found, interview_required
Dim xfs_screening, xfs_screening_display, caf_one_income, caf_one_assets, caf_one_resources, caf_one_rent, caf_one_utilities, caf_one_expenses, exp_det_case_note_found
Dim snap_exp_yn, snap_denial_date, interview_completed_case_note_found, interview_with, interview_type, verifications_requested_case_note_found, verifs_needed, caf_qualifying_questions_case_note_found
Dim verif_snap_checkbox, verif_cash_checkbox, verif_mfip_checkbox, verif_dwp_checkbox, verif_msa_checkbox, verif_ga_checkbox, verif_grh_checkbox, verif_emer_checkbox, verif_hc_checkbox
Dim qual_question_one, qual_memb_one, qual_question_two, qual_memb_two, qual_question_three, qual_memb_three, qual_question_four, qual_memb_four, qual_question_five, qual_memb_five, appt_notc_sent_on
Dim appt_date_in_note, addr_line_one, addr_line_two, city, state, zip, addr_county, homeless_yn, reservation_yn, addr_verif, living_situation, addr_eff_date, addr_future_date, mail_line_one
Dim mail_line_two, mail_city_line, mail_state_line, mail_zip_line, notes_on_address, total_shelter_amount, full_shelter_details, shelter_details, shelter_details_two, shelter_details_three
Dim prosp_heat_air, prosp_electric, prosp_phone, ABPS, ACCI, notes_on_acct, notes_on_acut, AREP, BILS, notes_on_cash, notes_on_cars, notes_on_coex, notes_on_dcex, DIET, DISA, EMPS
Dim FACI, FMED, IMIG, INSA, cit_id, other_assets, case_changes, PREG, earned_income, notes_on_rest, SCHL, notes_on_jobs, notes_on_time, notes_on_sanction, notes_on_wreg, full_abawd_info, notes_on_abawd
Dim notes_on_abawd_two, notes_on_abawd_three, programs_applied_for, TIKL_checkbox, interview_memb_list, shel_memb_list, verification_memb_list, notes_on_busi, Used_Interpreter_checkbox
Dim arep_id_info, CS_forms_sent_date, notes_on_ssa_income, notes_on_VA_income, notes_on_WC_income, other_uc_income_notes, notes_on_other_UNEA, hest_information, notes_on_other_deduction, expense_notes
Dim address_confirmation_checkbox, manual_total_shelter, app_month_assets, confirm_no_account_panel_checkbox, notes_on_other_assets, MEDI, DISQ, full_determination_done
Dim determined_income, determined_unea_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test
Dim is_elig_XFS, approval_date, applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, snap_denial_explain, case_assesment_text, next_steps_one, next_steps_two, next_steps_three
Dim next_steps_four, postponed_verifs_yn, list_postponed_verifs, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn
Dim other_state_verified_benefit_end_date, mn_elig_begin_date, action_due_to_out_of_state_benefits, case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done
Dim previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn, delay_action_due_to_faci
Dim deny_snap_due_to_faci, faci_review_completed, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn, next_er_month
Dim next_er_year, CAF_status, actions_taken, application_signed_checkbox, eDRS_sent_checkbox, updated_MMIS_checkbox, WF1_checkbox, Sent_arep_checkbox, intake_packet_checkbox, IAA_checkbox, recert_period_checkbox
Dim R_R_checkbox, E_and_T_checkbox, elig_req_explained_checkbox, benefit_payment_explained_checkbox, other_notes, client_delay_checkbox, client_delay_TIKL_checkbox, verif_req_form_sent_date
Dim adult_cash_count, child_cash_count, pregnant_caregiver_checkbox, adult_snap_count, child_snap_count, adult_emer_count, child_emer_count, EATS, relationship_detail, income_review_completed
Dim jobs_income_yn, busi_income_yn, unea_income_yn, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount
Dim tax_amount, room_amount, garage_amount, subsidy_amount, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities, do_we_have_applicant_id

full_determination_done = False
first_time_to_exp_det = True

HH_memb_row = 5 'This helps the navigation buttons work!
application_signed_checkbox = checked 'The script should default to having the application signed.
verifs_needed = "[Information here creates a SEPARATE CASE/NOTE.]"

member_count = 0
adult_cash_count = 0
child_cash_count = 0
adult_snap_count = 0
child_snap_count = 0
adult_emer_count = 0
child_emer_count = 0

'===========================================================================================================================

'FUNCTIONS =================================================================================================================
'===========================================================================================================================

'Specialty functionality
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_special_function = ucase(windows_user_ID)

'SCRIPT ====================================================================================================================
EMConnect ""
get_county_code				'since there is a county specific checkbox, this makes the the county clear
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
Call remove_dash_from_droplist(county_list)
Call find_user_name(worker_name)
script_run_lowdown = ""
case_notes_information = "No CASE NOTEs Attempted"
script_information_was_restored = False
CASH_checkbox = unchecked
GRH_checkbox = unchecked
SNAP_checkbox = unchecked
HC_checkbox = unchecked
EMER_checkbox = unchecked

allow_CASH_untrack = False
allow_GRH_untrack = False
allow_SNAP_untrack = False
allow_HC_untrack = True
allow_EMER_untrack = False

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 135, "CAF Script Case number dialog"
  EditBox 55, 80, 40, 15, MAXIS_case_number
  DropListBox 10, 115, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"HUF (DHS-8107)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"MNbenefits"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  ButtonGroup ButtonPressed
    OkButton 170, 115, 50, 15
    CancelButton 225, 115, 50, 15
  Text 5, 5, 270, 20, "This script is used to enter additional details about all information used to make eligibility determinations. "
  Text 5, 30, 215, 10, "STAT Panels should all be updated PRIOR to running this script."
  Text 5, 45, 275, 25, "This script should NOT be used to document information about the Interview. The script is set up to be a summary about processing a CAF (or similar) form after the interview. The correct fields to capture interview details are not present in this script."
  Text 5, 85, 50, 10, "Case number:"
  Text 10, 105, 90, 10, "Form Received in Agency:"
EndDialog
Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

        If CAF_form = "Select One:" then err_msg = err_msg & vbnewline & "* You must select the CAF form received."
        Call validate_MAXIS_case_number(err_msg, "*")
        IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If CAF_form = "CAF Addendum (DHS-5223C)" Then
    Call run_from_GitHub(script_repository & "notes/caf-addendum.vbs")
End If
If CAF_form = "MNbenefits" Then CAF_form = "CAF (DHS-5223) from MNbenefits"

developer_mode = False

Call back_to_SELF
continue_in_inquiry = ""
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
	If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
End If
If MX_region = "TRAINING" Then developer_mode = True

vars_filled = False
Call restore_your_work(vars_filled)			'looking for a 'restart' run

If vars_filled = False Then

	script_run_lowdown = script_run_lowdown & vbCr & "CAF Form: " & CAF_form
	processing_footer_month = ""
	processing_footer_year = ""

	Do
		Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
		If is_this_priv = True Then Call script_end_procedure("This case is PRIVILEGED and cannot be accessed. Request access to the case first and retry the script once you have access to the case.")
		EMReadScreen panel_prog_check, 4, 2, 50
	Loop until panel_prog_check = "PROG"
	EMReadScreen case_pw, 7, 21, 21

	snap_with_mfip = False
	MAXIS_footer_month = CM_mo
	MAXIS_footer_year = CM_yr
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, active_programs, programs_applied_for)
	EMReadScreen worker_id_for_data_table, 7, 21, 14
	EMReadScreen case_name_for_data_table, 25, 21, 40
	case_name_for_data_table = trim(case_name_for_data_table)

	CM_minus_1_mo =  right("0" &             DatePart("m",           DateAdd("m", -1, date)            ), 2)
	CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
	cash_terminated_revw_date = ""
	grh_terminated_revw_date = ""
	snap_terminated_revw_date = ""

	call date_array_generator(CM_minus_1_mo, CM_minus_1_yr, date_array)
	call Back_to_SELF
	For each month_date in date_array
		get_dates = False
		Call convert_date_into_MAXIS_footer_month(month_date, MAXIS_footer_month, MAXIS_footer_year)

		Call navigate_to_MAXIS_screen("STAT", "REVW")

		EMReadScreen cash_revw_code, 1, 7, 40
		EMReadScreen snap_revw_code, 1, 7, 60
		EMReadScreen hc_revw_code, 1, 7, 73
		If cash_revw_code = "N" or cash_revw_code = "U" or cash_revw_code = "I" or cash_revw_code = "A" or cash_revw_code = "T" Then
			get_dates = True
			If family_cash_case = True or adult_cash_case = True or grh_case = False Then
				the_process_for_cash = "Recertification"
				cash_recert_mo = MAXIS_footer_month
				cash_recert_yr = MAXIS_footer_year
				If cash_revw_code = "T" Then cash_terminated_revw_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year

				If (cash_revw_code = "I" or cash_revw_code = "A") and MAXIS_footer_month <> CM_plus_1_mo Then allow_CASH_untrack = True
			End If
			If grh_case = True Then
				the_review_is_ER = False
				EMReadScreen next_revw_process, 2, 9, 46
				grh_terminated_revw_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year
				If next_revw_process = "SR" Then the_review_is_ER = True
				If next_revw_process = "ER" Then
					Call write_value_and_transmit("X", 5, 35)
					EMReadScreen sr_date_month, 2, 9, 26
					If sr_date_month <> MAXIS_footer_month Then the_review_is_ER = True
					EMReadScreen er_date_month, 2, 9, 64
					If er_date_month = MAXIS_footer_month Then the_review_is_ER = True
					PF3
				End If
				If the_review_is_ER = True Then
					the_process_for_grh = "Recertification"
					grh_recert_mo = MAXIS_footer_month
					grh_recert_yr = MAXIS_footer_year
					If (cash_revw_code = "I" or cash_revw_code = "A") and MAXIS_footer_month <> CM_plus_1_mo Then allow_GRH_untrack = True
				End If
				' MsgBox "next_revw_process - " & next_revw_process & vbCr & "sr_date_month - " & sr_date_month & vbCr & "er_date_month - " & er_date_month & vbCr & "the_review_is_ER - " & the_review_is_ER & vbCr & "the_process_for_grh - " & the_process_for_grh
			End If
			If mfip_case = True Then snap_with_mfip = True
			If processing_footer_month = "" Then
				processing_footer_month = MAXIS_footer_month
				processing_footer_year = MAXIS_footer_year
			End If
		End If
		If snap_revw_code = "N" or snap_revw_code = "U" or snap_revw_code = "I" or snap_revw_code = "A" or snap_revw_code = "T" Then
			the_review_is_ER = False
			Call write_value_and_transmit("X", 5, 58)
			EMReadScreen er_date_month, 2, 9, 64
			If er_date_month = MAXIS_footer_month Then the_review_is_ER = True
			If snap_revw_code = "T" Then snap_terminated_revw_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year

			PF3

			If the_review_is_ER = True Then
				get_dates = True
				snap_with_mfip = False
				the_process_for_snap = "Recertification"
				snap_recert_mo = MAXIS_footer_month
				snap_recert_yr = MAXIS_footer_year
				If (snap_revw_code = "I" or snap_revw_code = "A") and MAXIS_footer_month <> CM_plus_1_mo Then allow_SNAP_untrack = True
				If processing_footer_month = "" Then
					processing_footer_month = MAXIS_footer_month
					processing_footer_year = MAXIS_footer_year
				End If
			End If
		End If
		'TODO - remove this? Why do we have it? How does this help us?
		' If hc_revw_code = "N" or hc_revw_code = "U" or hc_revw_code = "I" or hc_revw_code = "A" or hc_revw_code = "T" Then
		' 	get_dates = True
		' 	the_process_for_hc = "Recertification"
		' 	hc_recert_mo = MAXIS_footer_month
		' 	hc_recert_yr = MAXIS_footer_year
		' 	If processing_footer_month = "" Then
		' 		processing_footer_month = MAXIS_footer_month
		' 		processing_footer_year = MAXIS_footer_year
		' 	End If
		' End If

		If get_dates = True Then
			EMReadScreen REVW_CAF_datestamp, 8, 13, 37                       'reading theform date on REVW
			REVW_CAF_datestamp = replace(REVW_CAF_datestamp, " ", "/")
			If isdate(REVW_CAF_datestamp) = True then
				REVW_CAF_datestamp = cdate(REVW_CAF_datestamp) & ""
			Else
				REVW_CAF_datestamp = ""
			End if

			EMReadScreen REVW_interview_date, 8, 15, 37                       'reading the interview date on REVW
			REVW_interview_date = replace(REVW_interview_date, " ", "/")
			If isdate(REVW_interview_date) = True then
				REVW_interview_date = cdate(REVW_interview_date) & ""
			Else
				REVW_interview_date = ""
			End if

		End If
		' MsgBox "MAXIS_footer_month - " & MAXIS_footer_month & vbCr & "get_dates - " & get_dates & vbCr & "REVW_CAF_datestamp - " & REVW_CAF_datestamp & vbCr & "REVW_interview_date - " & REVW_interview_date

		IF SNAP_checkbox = checked THEN																															'checking for SNAP 24 month renewals.'
			EMWriteScreen "X", 05, 58																																	'opening the FS revw screen.
			transmit
			EMReadScreen SNAP_recert_date, 8, 9, 64
			PF3
			SNAP_recert_date = replace(SNAP_recert_date, " ", "/")
			If SNAP_recert_date <> "__/01/__" Then 																	'replacing the read blank spaces with / to make it a date
				SNAP_recert_compare_date = dateadd("m", "12", MAXIS_footer_month & "/01/" & MAXIS_footer_year)		'making a dummy variable to compare with, by adding 12 months to the requested footer month/year.
				IF datediff("d", SNAP_recert_compare_date, SNAP_recert_date) > 0 THEN											'If the read recert date is more than 0 days away from 12 months plus the MAXIS footer month/year then it is likely a 24 month period.'
					SNAP_recert_is_likely_24_months = TRUE
				ELSE
					SNAP_recert_is_likely_24_months = FALSE																									'otherwise if we don't we set it as false
				END IF
			Else
				SNAP_recert_is_likely_24_months = FALSE
			End If
		END IF

	Next

	If unknown_cash_pending = True Then the_process_for_cash = "Application"
	If ga_status = "PENDING" Then the_process_for_cash = "Application"
	If msa_status = "PENDING" Then the_process_for_cash = "Application"
	If mfip_status = "PENDING" Then the_process_for_cash = "Application"
	If dwp_status = "PENDING" Then the_process_for_cash = "Application"
	If grh_status = "PENDING" Then the_process_for_grh = "Application"
	If snap_status = "PENDING" Then the_process_for_snap = "Application"

	If the_process_for_cash <> "" Then CASH_checkbox = checked
	If the_process_for_grh <> "" Then GRH_checkbox = checked
	If the_process_for_snap <> "" Then SNAP_checkbox = checked
	If emer_status = "PENDING" Then
		EMER_checkbox = checked
		the_process_for_emer = "Application"
	End If
	If cash_terminated_revw_date <> "" then cash_terminated_revw_date = DateAdd("d", 0, cash_terminated_revw_date)
	If grh_terminated_revw_date <> "" then grh_terminated_revw_date = DateAdd("d", 0, grh_terminated_revw_date)
	If snap_terminated_revw_date <> "" then snap_terminated_revw_date = DateAdd("d", 0, snap_terminated_revw_date)

	If (CASH_checkbox = unchecked and GRH_checkbox = unchecked and SNAP_checkbox = unchecked and EMER_checkbox = unchecked) or IsDate(cash_terminated_revw_date) = True or IsDate(grh_terminated_revw_date) = True or IsDate(snap_terminated_revw_date) = True Then
		past_90_days = DateAdd("d", -90, date)

		Call navigate_to_MAXIS_screen("STAT", "PROG")
		EMReadScreen prog_cash_1_appl_date, 8, 6, 33
		EMReadScreen prog_cash_1_intv_date, 8, 6, 55
		EMReadScreen prog_cash_1_prog, 2, 6, 67
		EMReadScreen prog_cash_1_status, 4, 6, 74

		EMReadScreen prog_cash_2_appl_date, 8, 7, 33
		EMReadScreen prog_cash_2_intv_date, 8, 7, 55
		EMReadScreen prog_cash_2_prog, 2, 7, 67
		EMReadScreen prog_cash_2_status, 4, 7, 74

		EMReadScreen prog_emer_appl_date, 8, 8, 33
		EMReadScreen prog_emer_intv_date, 8, 8, 55
		EMReadScreen prog_emer_status, 4, 8, 74

		EMReadScreen prog_grh_appl_date, 8, 9, 33
		EMReadScreen prog_grh_intv_date, 8, 9, 55
		EMReadScreen prog_grh_status, 4, 9, 74

		EMReadScreen prog_snap_appl_date, 8, 10, 33
		EMReadScreen prog_snap_intv_date, 8, 10, 55
		EMReadScreen prog_snap_status, 4, 10, 74

		prog_cash_1_appl_date = replace(prog_cash_1_appl_date, " ", "/")
		If IsDate(prog_cash_1_appl_date) = True Then
			If IsDate(cash_terminated_revw_date) = True Then
				If DateDiff("d", cash_terminated_revw_date, prog_cash_1_appl_date) > 0 Then
					CASH_checkbox = checked
					allow_CASH_untrack = True
					the_process_for_cash = "Application"
					PROG_CAF_datestamp = prog_cash_1_appl_date
					prog_cash_1_intv_date = replace(prog_cash_1_intv_date, " ", "/")
					If prog_cash_1_intv_date <> "__/__/__" Then PROG_interview_date = prog_cash_1_intv_date
					cash_recert_mo = ""
					cash_recert_yr = ""
				End If
			ElseIf DateDiff("d", past_90_days, prog_cash_1_appl_date) >= 0 Then
				CASH_checkbox = checked
				allow_CASH_untrack = True
				the_process_for_cash = "Application"
				PROG_CAF_datestamp = prog_cash_1_appl_date
				prog_cash_1_intv_date = replace(prog_cash_1_intv_date, " ", "/")
				If prog_cash_1_intv_date <> "__/__/__" Then PROG_interview_date = prog_cash_1_intv_date
			End If
		End If
		prog_cash_2_appl_date = replace(prog_cash_2_appl_date, " ", "/")
		If IsDate(prog_cash_2_appl_date) = True Then

			If IsDate(cash_terminated_revw_date) = True Then
				If DateDiff("d", cash_terminated_revw_date, prog_cash_2_appl_date) > 0 Then
					CASH_checkbox = checked
					allow_CASH_untrack = True
					the_process_for_cash = "Application"
					PROG_CAF_datestamp = prog_cash_2_appl_date
					prog_cash_2_intv_date = replace(prog_cash_2_intv_date, " ", "/")
					If prog_cash_2_intv_date <> "__/__/__" Then PROG_interview_date = prog_cash_2_intv_date
					cash_recert_mo = ""
					cash_recert_yr = ""
				End If
			ElseIf DateDiff("d", past_90_days, prog_cash_2_appl_date) >= 0 Then
				CASH_checkbox = checked
				allow_CASH_untrack = True
				the_process_for_cash = "Application"
				PROG_CAF_datestamp = prog_cash_2_appl_date
				prog_cash_2_intv_date = replace(prog_cash_2_intv_date, " ", "/")
				If prog_cash_2_intv_date <> "__/__/__" Then PROG_interview_date = prog_cash_2_intv_date
			End If
		End If

		prog_emer_appl_date = replace(prog_emer_appl_date, " ", "/")
		If IsDate(prog_emer_appl_date) = True Then
			If DateDiff("d", past_90_days, prog_emer_appl_date) >= 0 Then
				EMER_checkbox = checked
				allow_EMER_untrack = True
				the_process_for_emer = "Application"
				PROG_CAF_datestamp = prog_emer_appl_date
				prog_emer_intv_date = replace(prog_emer_intv_date, " ", "/")
				If prog_emer_intv_date <> "__/__/__" Then PROG_interview_date = prog_emer_intv_date
			End If
		End If

		prog_grh_appl_date = replace(prog_grh_appl_date, " ", "/")
		If IsDate(prog_grh_appl_date) = True Then
			If IsDate(grh_terminated_revw_date) = True Then
				If DateDiff("d", grh_terminated_revw_date, prog_grh_appl_date) > 0 Then
					GRH_checkbox = checked
					allow_GRH_untrack = True
					the_process_for_grh = "Application"
					PROG_CAF_datestamp = prog_grh_appl_date
					prog_grh_intv_date = replace(prog_grh_intv_date, " ", "/")
					If prog_grh_intv_date <> "__/__/__" Then PROG_interview_date = prog_grh_intv_date
					grh_recert_mo = ""
					grh_recert_yr = ""
				End If
			ElseIf DateDiff("d", past_90_days, prog_grh_appl_date) >= 0 Then
				GRH_checkbox = checked
				allow_GRH_untrack = True
				the_process_for_grh = "Application"
				PROG_CAF_datestamp = prog_grh_appl_date
				prog_grh_intv_date = replace(prog_grh_intv_date, " ", "/")
				If prog_grh_intv_date <> "__/__/__" Then PROG_interview_date = prog_grh_intv_date
			End If
		End If

		prog_snap_appl_date = replace(prog_snap_appl_date, " ", "/")
		If IsDate(prog_snap_appl_date) = True Then
			If IsDate(snap_terminated_revw_date) = True Then
				If DateDiff("d", snap_terminated_revw_date, prog_snap_appl_date) > 0 Then
					SNAP_checkbox = checked
					allow_SNAP_untrack = True
					the_process_for_SNAP = "Application"
					PROG_CAF_datestamp = prog_snap_appl_date
					prog_snap_intv_date = replace(prog_snap_intv_date, " ", "/")
					If prog_snap_intv_date <> "__/__/__" Then PROG_interview_date = prog_snap_intv_date
					snap_recert_mo = ""
					snap_recert_yr = ""

				End If
			ElseIf DateDiff("d", past_90_days, prog_snap_appl_date) >= 0 Then
				SNAP_checkbox = checked
				allow_SNAP_untrack = True
				the_process_for_SNAP = "Application"
				PROG_CAF_datestamp = prog_snap_appl_date
				prog_snap_intv_date = replace(prog_snap_intv_date, " ", "/")
				If prog_snap_intv_date <> "__/__/__" Then PROG_interview_date = prog_snap_intv_date

			End If
		End If

		allow_SNAP_untrack = True

	End If

	If CASH_checkbox = unchecked and GRH_checkbox = unchecked and SNAP_checkbox = unchecked and EMER_checkbox = unchecked Then
		end_early_mgs = "This script (NOTES - CAF) could not find a program that was pending or coded as requiring a review."
		end_early_mgs = end_early_mgs & vbCr & vbCr & "The script could not find details about the programs being processed. The script checked REVW for the months " & CM_minus_1_mo & "/" & CM_minus_1_yr & " through " & CM_plus_1_mo & "/" & CM_plus_1_yr & " and checked CASE/CURR for pending programs."
		end_early_mgs = end_early_mgs & vbCr & vbCr & "If a CAF (or similar form) has been received, but an ER is not due and no additional programs are requested, review the case to determine the next processing steps to take."
		end_early_mgs = end_early_mgs & vbCr & vbCr & "The script 'NOTES - CAF' is to support the process typically associated with receipt of a CAF and not to be a comprehensive case review tool."

		Call script_end_procedure_with_error_report(end_early_mgs)
	End If

	MAXIS_footer_month = CM_mo
	MAXIS_footer_year = CM_yr

	If unknown_cash_pending = True or ga_status = "PENDING" or msa_status = "PENDING" or mfip_status = "PENDING" or dwp_status = "PENDING" or grh_status = "PENDING" or snap_status = "PENDING" or emer_status = "PENDING" Then
		Call navigate_to_MAXIS_screen("STAT", "PROG")

		'here we are going to read for the CAF date by reading each line of PROG and looking for the most recent date.
		row = 6
		Do
			EMReadScreen appl_prog_date, 8, row, 33
			If appl_prog_date <> "__ __ __" then appl_prog_date_array = appl_prog_date_array & replace(appl_prog_date, " ", "/") & " "

			row = row + 1
		Loop until row = 12
		appl_prog_date_array = split(appl_prog_date_array)
		PROG_CAF_datestamp = CDate(appl_prog_date_array(0))
		for i = 0 to ubound(appl_prog_date_array) - 1
			if CDate(appl_prog_date_array(i)) > PROG_CAF_datestamp then
				PROG_CAF_datestamp = CDate(appl_prog_date_array(i))
			End if
		next
		If isdate(PROG_CAF_datestamp) = True then
			var_month = datepart("m", PROG_CAF_datestamp)
			IF len(var_month) = 1 THEN var_month = "0" & var_month
			var_day = datepart("d", PROG_CAF_datestamp)
			IF len(var_day) = 1 THEN var_day = "0" & var_day
			var_year = right(datepart("yyyy", PROG_CAF_datestamp), 2)
			CAF_MAXIS_date = var_month & " " & var_day & " " & var_year
			PROG_CAF_datestamp = cdate(PROG_CAF_datestamp) & ""
		Else
			PROG_CAF_datestamp = ""
		End if

		cash_interview_missing = False          'defaulting to the interview date is NOT missing
		emer_interview_missing = False
		snap_interview_missing = False
		'checking Cash lines - which included GRH
		If cash_checkbox = checked Then
			EMReadScreen prog_cash_1_form_date, 8, 6, 33
			If prog_cash_1_form_date = CAF_MAXIS_date Then
				EMReadScreen prog_cash_1_intvw_date, 8, 6, 55
				cash_interview_missing = True
				If prog_cash_1_intvw_date <> "__ __ __" AND prog_cash_1_intvw_date <> "        " then
					PROG_interview_date = replace(prog_cash_1_intvw_date, " ", "/") & " "
					cash_interview_missing = False
				End If
			End If
			EMReadScreen prog_cash_2_form_date, 8, 7, 33
			If prog_cash_2_form_date = CAF_MAXIS_date Then
				EMReadScreen prog_cash_2_intvw_date, 8, 7, 55
				cash_interview_missing = True
				If prog_cash_2_intvw_date <> "__ __ __" AND prog_cash_2_intvw_date <> "        " then
					PROG_interview_date = replace(prog_cash_2_intvw_date, " ", "/") & " "
					cash_interview_missing = False
				End If
			End If
		End If
		If GRH_checkbox = checked Then
			EMReadScreen prog_grh_form_date, 8, 9, 33
			If prog_grh_form_date = CAF_MAXIS_date Then
				EMReadScreen prog_grh_intvw_date, 8, 9, 55
				cash_interview_missing = True
				If prog_grh_intvw_date <> "__ __ __" AND prog_grh_intvw_date <> "        " then
					PROG_interview_date = replace(prog_grh_intvw_date, " ", "/") & " "
					cash_interview_missing = False
				End If
			End If
		End If
		'checking EMER lines
		If EMER_checkbox = checked Then
			EMReadScreen prog_emer_form_date, 8, 8, 33
			If prog_emer_form_date = CAF_MAXIS_date Then
				EMReadScreen prog_emer_intvw_date, 8, 8, 55
				emer_interview_missing = True
				If prog_emer_intvw_date <> "__ __ __" AND prog_emer_intvw_date <> "        " then
					PROG_interview_date = replace(prog_emer_intvw_date, " ", "/") & " "
					emer_interview_missing = False
				End If
			End If
		End If
		'Checking SNAP lines
		If SNAP_checkbox = checked Then
			EMReadScreen prog_snap_form_date, 8, 10, 33
			If prog_snap_form_date = CAF_MAXIS_date Then
				EMReadScreen prog_snap_intvw_date, 8, 10, 55
				snap_interview_missing = True
				If prog_snap_intvw_date <> "__ __ __" AND prog_snap_intvw_date <> "        " then
					PROG_interview_date = replace(prog_snap_intvw_date, " ", "/") & " "
					snap_interview_missing = False
				End If
			End If
		End If
		'If any interview dates are missing we blank out the interview date variable
		If cash_interview_missing = True Then PROG_interview_date = ""
		If emer_interview_missing = True Then PROG_interview_date = ""
		If snap_interview_missing = True Then PROG_interview_date = ""
	End If

	multiple_CAF_dates = False
	multiple_interview_dates = False
	' MsgBox "REVW_CAF_datestamp - " & REVW_CAF_datestamp & vbCr & "PROG_CAF_datestamp - " & PROG_CAF_datestamp & vbCr & "CAF_datestamp - " & CAF_datestamp & vbCr & "1"
	If PROG_CAF_datestamp <> "" And REVW_CAF_datestamp <> "" and PROG_CAF_datestamp <> REVW_CAF_datestamp Then
		CAF_datestamp = REVW_CAF_datestamp
		If DateDiff("d", PROG_CAF_datestamp, REVW_CAF_datestamp) Then CAF_datestamp = PROG_CAF_datestamp
		' MsgBox "REVW_CAF_datestamp - " & REVW_CAF_datestamp & vbCr & "PROG_CAF_datestamp - " & PROG_CAF_datestamp & vbCr & "CAF_datestamp - " & CAF_datestamp & vbCr & "2"
		multiple_CAF_dates = True
	Else
		If PROG_CAF_datestamp <> "" Then CAF_datestamp = PROG_CAF_datestamp
		If REVW_CAF_datestamp <> "" Then CAF_datestamp = REVW_CAF_datestamp
		' MsgBox "REVW_CAF_datestamp - " & REVW_CAF_datestamp & vbCr & "PROG_CAF_datestamp - " & PROG_CAF_datestamp & vbCr & "CAF_datestamp - " & CAF_datestamp & vbCr & "3"
	End If
	' MsgBox "REVW_CAF_datestamp - " & REVW_CAF_datestamp & vbCr & "PROG_CAF_datestamp - " & PROG_CAF_datestamp & vbCr & "CAF_datestamp - " & CAF_datestamp & vbCr & "4"

	If IsDate(PROG_interview_date) = True Then Call convert_date_into_MAXIS_footer_month(PROG_interview_date, processing_footer_month, processing_footer_year)
	If PROG_interview_date <> "" And REVW_interview_date <> "" and PROG_interview_date <> REVW_interview_date Then
		interview_date = REVW_interview_date
		If DateDiff("d", PROG_interview_date, REVW_interview_date) Then interview_date = PROG_interview_date
		multiple_interview_dates = True
	Else
		If PROG_interview_date <> "" Then interview_date = PROG_interview_date
		If REVW_interview_date <> "" Then interview_date = REVW_interview_date
	End If

	If IsDate(CAF_datestamp) = True and interview_date = "" Then

		Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
		too_old_date = DateAdd("D", -1, CAF_datestamp)              'We don't need to read notes from before the CAF date

		note_row = 5
		Do
			EMReadScreen note_date, 8, note_row, 6                  'reading the note date

			EMReadScreen note_title, 55, note_row, 25               'reading the note header
			note_title = trim(note_title)

			'INTERVIEW CNOTE
			If left(note_title, 24) = "~ Interview Completed on" Then
				interview_completed_case_note_found = True

				EMWriteScreen "X", note_row, 3                          'Opening the Interview Note to read some interview details
				transmit

				EMReadScreen in_correct_note, 24, 4, 3                  'making sure we are in the right note
				EMReadScreen note_list_header, 23, 4, 25

				If in_correct_note = "~ Interview Completed on" Then
					in_note_row = 5
					Do
						EMReadScreen first_part_of_line, 12, in_note_row, 3                         'Reading the header portion
						EMReadScreen whole_note_line, 77, in_note_row, 3                            'Reading all the line information
						whole_note_line = trim(whole_note_line)
						If first_part_of_line = "Completed wi" Then                                 'COMPLETED WITH header has person and type information
							whole_note_line = replace(whole_note_line, "Completed with ", "")       'removes the header
							position = Instr(whole_note_line, " via")                               'finds the dividing point in the content which is always the word 'via'
							with_len = position + 4

							interview_with = left(whole_note_line, position)                        'reading the person that did the interview - which is to the left of the dividing point
							interview_type = right(whole_note_line, len(whole_note_line) - with_len)'the type is anything that is NOT the with
							interview_with = trim(interview_with)
							interview_type = trim(interview_type)
						End If
						If first_part_of_line = "Completed on" Then                                 'COMPLETED ON header has interview date
							whole_note_line = replace(whole_note_line, "Completed on ", "")         'removes the header
							position = Instr(whole_note_line, " at")                                'finds the dividing point

							interview_date = left(whole_note_line, position)                        'interview date is to the left of the dividing point'
							interview_date = trim(interview_date)
						End If
						in_note_row = in_note_row + 1
						If interview_with <> "" AND interview_type <> "" AND interview_date <> "" Then Exit Do      'if we found all of it, we can be done
					Loop until trim(whole_note_line) = ""
					PF3         'leaving the note.

				Else
					If note_list_header <> "First line of Case note" Then PF3           'this backs us out of the note if we ended up in the wrong note.
				End If
			End If

			if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

			note_row = note_row + 1
			if note_row = 19 then
				note_row = 5
				PF8
				EMReadScreen check_for_last_page, 9, 24, 14
				If check_for_last_page = "LAST PAGE" Then Exit Do
			End If
			EMReadScreen next_note_date, 8, note_row, 6
			if next_note_date = "        " then Exit Do
		Loop until DateDiff("d", too_old_date, next_note_date) <= 0
	End If

	MAXIS_footer_month = processing_footer_month
	MAXIS_footer_year =  processing_footer_year

	script_run_lowdown = script_run_lowdown & vbCr & vbCr & "determine_program_and_case_status_from_CASE_CURR"
	script_run_lowdown = script_run_lowdown & vbCr & "case_active - " & case_active
	script_run_lowdown = script_run_lowdown & vbCr & "case_pending - " & case_pending
	script_run_lowdown = script_run_lowdown & vbCr & "case_rein - " & case_rein
	script_run_lowdown = script_run_lowdown & vbCr & "family_cash_case - " & family_cash_case
	script_run_lowdown = script_run_lowdown & vbCr & "mfip_case - " & mfip_case
	script_run_lowdown = script_run_lowdown & vbCr & "dwp_case - " & dwp_case
	script_run_lowdown = script_run_lowdown & vbCr & "adult_cash_case - " & adult_cash_case
	script_run_lowdown = script_run_lowdown & vbCr & "ga_case - " & ga_case
	script_run_lowdown = script_run_lowdown & vbCr & "msa_case - " & msa_case
	script_run_lowdown = script_run_lowdown & vbCr & "grh_case - " & grh_case
	script_run_lowdown = script_run_lowdown & vbCr & "snap_case - " & snap_case
	script_run_lowdown = script_run_lowdown & vbCr & "ma_case - " & ma_case
	script_run_lowdown = script_run_lowdown & vbCr & "msp_case - " & msp_case
	script_run_lowdown = script_run_lowdown & vbCr & "emer_case - " & emer_case
	script_run_lowdown = script_run_lowdown & vbCr & "unknown_cash_pending - " & unknown_cash_pending
	script_run_lowdown = script_run_lowdown & vbCr & "unknown_hc_pending - " & unknown_hc_pending
	script_run_lowdown = script_run_lowdown & vbCr & "ga_status - " & ga_status
	script_run_lowdown = script_run_lowdown & vbCr & "msa_status - " & msa_status

	script_run_lowdown = script_run_lowdown & vbCr & "mfip_status - " & mfip_status
	script_run_lowdown = script_run_lowdown & vbCr & "dwp_status - " & dwp_status
	script_run_lowdown = script_run_lowdown & vbCr & "grh_status - " & grh_status
	script_run_lowdown = script_run_lowdown & vbCr & "snap_status - " & snap_status
	script_run_lowdown = script_run_lowdown & vbCr & "ma_status - " & ma_status
	script_run_lowdown = script_run_lowdown & vbCr & "msp_status - " & msp_status
	script_run_lowdown = script_run_lowdown & vbCr & "msp_type - " & msp_type
	script_run_lowdown = script_run_lowdown & vbCr & "emer_status - " & emer_status

	script_run_lowdown = script_run_lowdown & vbCr & "emer_type - " & emer_type
	script_run_lowdown = script_run_lowdown & vbCr & "case_status - " & case_status
	script_run_lowdown = script_run_lowdown & vbCr & "active_programs - " & active_programs
	script_run_lowdown = script_run_lowdown & vbCr & "programs_applied_for - " & programs_applied_for

	script_run_lowdown = script_run_lowdown & vbCr & vbCr & "Reading of DATES from STAT"
	script_run_lowdown = script_run_lowdown & vbCr & "REVW_CAF_datestamp - " & REVW_CAF_datestamp
	script_run_lowdown = script_run_lowdown & vbCr & "REVW_interview_date - " & REVW_interview_date
	script_run_lowdown = script_run_lowdown & vbCr & "PROG_CAF_datestamp - " & PROG_CAF_datestamp
	script_run_lowdown = script_run_lowdown & vbCr & "PROG_interview_date - " & PROG_interview_date
	script_run_lowdown = script_run_lowdown & vbCr & "CAF_datestamp - " & CAF_datestamp
	script_run_lowdown = script_run_lowdown & vbCr & "interview_date - " & interview_date
	script_run_lowdown = script_run_lowdown & vbCr & "multiple_CAF_dates - " & multiple_CAF_dates
	script_run_lowdown = script_run_lowdown & vbCr & "multiple_interview_dates - " & multiple_interview_dates


	script_run_lowdown = script_run_lowdown & vbCr & vbCr & "Script determination of programs"
	If CASH_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH checkbox - CHECKED"
	If CASH_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "type_of_cash - " & type_of_cash
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_cash - " & the_process_for_cash
	script_run_lowdown = script_run_lowdown & vbCr & "cash_recert_mo - " & cash_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "cash_recert_yr - " & cash_recert_yr
	If GRH_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "GRH checkbox - CHECKED"
	If GRH_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "GRH checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_grh - " & the_process_for_grh
	script_run_lowdown = script_run_lowdown & vbCr & "grh_recert_mo - " & grh_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "grh_recert_yr - " & grh_recert_yr
	If SNAP_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP checkbox - CHECKED"
	If SNAP_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "snap_with_mfip - " & snap_with_mfip
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_snap - " & the_process_for_snap
	script_run_lowdown = script_run_lowdown & vbCr & "snap_recert_mo - " & snap_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "snap_recert_yr - " & snap_recert_yr
	If HC_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "HC checkbox - CHECKED"
	If HC_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "HC checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_hc - " & the_process_for_hc
	script_run_lowdown = script_run_lowdown & vbCr & "hc_recert_mo - " & hc_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "hc_recert_yr - " & hc_recert_yr
	If EMER_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER checkbox - CHECKED"
	If EMER_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "type_of_emer - " & type_of_emer
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_emer - " & the_process_for_emer
	script_run_lowdown = script_run_lowdown & vbCr & "Footer month: " & MAXIS_footer_month & "/" & MAXIS_footer_year

	If CASH_checkbox = unchecked Then allow_CASH_untrack = True
	If SNAP_checkbox = unchecked Then allow_SNAP_untrack = True
	If EMER_checkbox = unchecked Then allow_EMER_untrack = True
	If GRH_checkbox = unchecked Then allow_GRH_untrack = True

	option_to_process_with_no_interview = False
	original_snap_with_mfip = snap_with_mfip
	Do
		DO
			err_msg = ""
			If the_process_for_cash = "Recertification" and adult_cash_case = True Then option_to_process_with_no_interview = True
			If the_process_for_grh = "Recertification" Then option_to_process_with_no_interview = True
			If interview_date <> "" Then option_to_process_with_no_interview = False
			' MsgBox "the_process_for_grh - " & the_process_for_grh & vbCr & "option_to_process_with_no_interview - " & option_to_process_with_no_interview & vbCr & "interview_date - " & interview_date

			' MsgBox "adult_cash_case - " & adult_cash_case & vbCr & "family_cash_case - "& family_cash_case
			If adult_cash_case = TRUE Then type_of_cash = "Adult"
			If family_cash_case = TRUE Then type_of_cash = "Family"
			dlg_len = 205
			y_pos = 60
			dlg_wdth = 275

			If multiple_CAF_dates = True or multiple_interview_dates = True Then
				PROG_CAF_Form = CAF_Form
				REVW_CAF_Form = CAF_Form
				If PROG_CAF_Form = "CAF (DHS-5223) from MNbenefits" then PROG_CAF_Form = "MNbenefits"
				If REVW_CAF_Form = "CAF (DHS-5223) from MNbenefits" then REVW_CAF_Form = "MNbenefits"

				dlg_wdth = 365
				dlg_len = dlg_len + 60
			End If
			If option_to_process_with_no_interview = True Then
				dlg_len = dlg_len + 10
				dlg_wdth = 365
			End if
			If snap_with_mfip = True Then dlg_len = dlg_len + 15

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, dlg_wdth, dlg_len, "CAF Process"
			ButtonGroup ButtonPressed
				Text 10, 10, 100, 10, "CAF Date:" & CAF_datestamp
				If multiple_CAF_dates = True or multiple_interview_dates = True Then Text 110, 10, 195, 10, "(If multiple dates found, the oldest date is used)"

				If option_to_process_with_no_interview = True Then
					Text 10, 20, 315, 10, "No Interview date entered, but this case appears to be possibly MSA/GA/GRH at Recertification."
					CheckBox 10, 30, 250, 10, "Check here to process ONLY Adult cash at ER without interview", adult_cash_er_no_interview_checkbox
					y_pos = 45
				Else
					If interview_date <> "" Then Text 10, 20, 100, 10, "Interview Date:" & interview_date
					If interview_date = "" Then Text 10, 20, 100, 10, "NO Interview Date"
					y_pos = 35
				End If
				Text 10, y_pos, 200, 10, "Program(s) Requiring review and determinations:"
				Text 10, y_pos+15, 35, 10, "Program"
				Text 85, y_pos+15, 65, 10, "Eligibility Process"
				Text 160, y_pos+15, 50, 10, "Recert MM/YY"
				y_pos = y_pos + 25

				If CASH_checkbox = checked Then
					Text 15, y_pos + 5, 20, 10, "Cash"
					DropListBox 35, y_pos, 45, 45, "Select"+chr(9)+"Family"+chr(9)+"Adult", type_of_cash
					DropListBox 85, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", the_process_for_cash
					EditBox 160, y_pos, 20, 15, cash_recert_mo
					EditBox 185, y_pos, 20, 15, cash_recert_yr
					If allow_CASH_untrack = True Then PushButton 215, y_pos, 60, 13, "Untrack CASH", untrack_cash_btn
					y_pos = y_pos + 20
				End If

				If GRH_checkbox = checked Then
					Text 15, y_pos + 5, 20, 10, "GRH"
					DropListBox 85, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", the_process_for_grh
					EditBox 160, y_pos, 20, 15, grh_recert_mo
					EditBox 185, y_pos, 20, 15, grh_recert_yr
					If allow_GRH_untrack = True Then PushButton 215, y_pos, 60, 13, "Untrack GRH", untrack_grh_btn
					y_pos = y_pos + 20
				End If
				If snap_with_mfip = True Then
					Text 20, y_pos, 200, 10, "SNAP benefit is a part of MFIP Grant"
					y_pos = y_pos + 15
				End If
				If SNAP_checkbox = checked Then
					If snap_with_mfip = False Then Text 15, y_pos + 5, 20, 10, "SNAP"
					DropListBox 85, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification", the_process_for_snap
					EditBox 160, y_pos, 20, 15, snap_recert_mo
					EditBox 185, y_pos, 20, 15, snap_recert_yr
					If allow_SNAP_untrack = True Then PushButton 215, y_pos, 60, 13, "Untrack SNAP", untrack_snap_btn
					y_pos = y_pos + 20
				End If
				If HC_checkbox = checked Then
					Text 15, y_pos + 5, 40, 10, "Health Care"
					DropListBox 85, y_pos, 65, 45, "Select One..."+chr(9)+"Recertification", the_process_for_hc
					EditBox 160, y_pos, 20, 15, hc_recert_mo
					EditBox 185, y_pos, 20, 15, hc_recert_yr
					If allow_HC_untrack = True Then PushButton 215, y_pos, 60, 13, "Untrack HC", untrack_hc_btn
					y_pos = y_pos + 20
				End If
				If EMER_checkbox = checked Then
					Text 15, y_pos+5, 40, 10, "EMER"
					DropListBox 35, y_pos, 45, 45, "Select"+chr(9)+"EA"+chr(9)+"EGA", type_of_emer
					DropListBox 85, y_pos, 65, 45, "Select One..."+chr(9)+"Application", the_process_for_emer
					If allow_EMER_untrack = True Then PushButton 215, y_pos, 60, 13, "Untrack EMER", untrack_emer_btn
					y_pos = y_pos + 20
				End If

				Text 10, y_pos, 210, 10,  "Add any programs to review that are not coded in STAT or REVW:"
				y_pos = y_pos + 15
				If CASH_checkbox = unchecked Then
					PushButton 20, y_pos, 65, 13, "Add CASH", check_cash_box_btn
					y_pos = y_pos + 20
				End If
				If GRH_checkbox = unchecked Then
					PushButton 20, y_pos, 65, 13, "Add GRH", check_grh_box_btn
					y_pos = y_pos + 20
				End If
				If SNAP_checkbox = unchecked Then
					PushButton 20, y_pos, 65, 13, "Add SNAP", check_snap_box_btn
					y_pos = y_pos + 20
				End If
				If HC_checkbox = unchecked Then
					PushButton 20, y_pos, 65, 13, "Add HC", check_hc_box_btn
					y_pos = y_pos + 20
				End If
				If EMER_checkbox = unchecked Then
					PushButton 20, y_pos, 65, 13, "Add EMER", check_emer_box_btn
					y_pos = y_pos + 20
				End If

				If multiple_CAF_dates = True or multiple_interview_dates = True Then
					y_pos = y_pos + 10
					Text 10, y_pos, 300, 20, "This case has different Form Dates and/or Interview Dates entered. Document the form for each of these dates/interview"
					y_pos = y_pos + 20
					Text 10, y_pos+5, 100, 10, "REVW CAF Date: " & REVW_CAF_datestamp
					Text 115, y_pos+5, 100, 10, "REVW Interview: " & REVW_interview_date
					DropListBox 220, y_pos, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"HUF (DHS-8107)"+chr(9)+"MNbenefits"+chr(9)+"Combined AR for Certain Pops (DHS-3727)", REVW_CAF_Form
					y_pos = y_pos + 20
					Text 10, y_pos+5, 100, 10, "PROG CAF Date: " & PROG_CAF_datestamp
					Text 115, y_pos+5, 100, 10, "PROG Interview: " & PROG_interview_date
					DropListBox 220, y_pos, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"MNbenefits", PROG_CAF_Form
					y_pos = y_pos + 20
				End If
				Text 10, dlg_len-15, 50, 10, "Footer Month: "
				EditBox 60, dlg_len-20, 20, 15, MAXIS_footer_month
				EditBox 85, dlg_len-20, 20, 15, MAXIS_footer_year
				' y_pos = y_pos + 5
				OkButton dlg_wdth-55, dlg_len-20, 50, 15
			EndDialog


			Dialog Dialog1
			cancel_confirmation

			If len(cash_recert_yr) = 4 AND left(cash_recert_yr, 2) = "20" Then cash_recert_yr = right(cash_recert_yr, 2)
			If len(grh_recert_yr) = 4 AND left(grh_recert_yr, 2) = "20" Then grh_recert_yr = right(grh_recert_yr, 2)
			If len(snap_recert_yr) = 4 AND left(snap_recert_yr, 2) = "20" Then snap_recert_yr = right(snap_recert_yr, 2)
			If len(hc_recert_yr) = 4 AND left(hc_recert_yr, 2) = "20" Then hc_recert_yr = right(hc_recert_yr, 2)

			Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
			If CASH_checkbox = checked Then
				If type_of_cash = "Select" Then err_msg = err_msg & vbNewLine & "* Indicate if the cash program is a family or adult cash request."
				If the_process_for_cash = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the CASH program is at application or recertification."
				If the_process_for_cash = "Recertification" AND (len(cash_recert_mo) <> 2 or len(cash_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For CASH at recertification, enter the footer month and year the of the recertification."
			End If
			If GRH_checkbox = checked Then
				If the_process_for_grh = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the CASH program is at application or recertification."
				If the_process_for_grh = "Recertification" AND (len(grh_recert_mo) <> 2 or len(grh_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For GRH at recertification, enter the footer month and year the of the recertification."
			End If
			If SNAP_checkbox = checked Then
				If the_process_for_snap = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the SNAP program is at application or recertification."
				If the_process_for_snap = "Recertification" AND (len(snap_recert_mo) <> 2 or len(snap_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For SNAP at recertification, enter the footer month and year the of the recertification."
			End If
			If HC_checkbox = checked Then
				If the_process_for_hc = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the Health Care program is at application or recertification."
				If the_process_for_hc = "Recertification" AND (len(hc_recert_mo) <> 2 or len(hc_recert_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For HC at recertification, enter the footer month and year the of the recertification."
				If CASH_checkbox = unchecked and GRH_checkbox = unchecked and SNAP_checkbox = unchecked and EMER_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* HC cannot be processed independently in this script, either select another program or cancel this script run and run a different script to support HC only."
			End If
			If EMER_checkbox = checked Then
				If the_process_for_hc = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the Emergency program is at application or recertification."
			End If
			If adult_cash_er_no_interview_checkbox = checked Then
				If CASH_checkbox = checked Then
					If the_process_for_cash <> "Recertification" Then err_msg = err_msg & vbNewLine & "* If you have selected to process adult cash cases with no interview needed and no interview completed, the cash process should be recertification."
				End If
				If GRH_checkbox = checked Then
					If the_process_for_grh <> "Recertification" Then err_msg = err_msg & vbNewLine & "* If you have selected to process adult cash cases with no interview needed and no interview completed, the grh process should be recertification."
				End If
				SNAP_checkbox = unchecked
				HC_checkbox = unchecked
				EMER_checkbox = unchecked
			End If
			If multiple_CAF_dates = True or multiple_interview_dates = True Then
				If REVW_CAF_Form = "Select One:" Then err_msg = err_msg & vbNewLine & "* Enter the form that was provided on " & REVW_CAF_datestamp & " and entered into the REVW panel."
				If PROG_CAF_Form = "Select One:" Then err_msg = err_msg & vbNewLine & "* Enter the form that was provided on " & PROG_CAF_datestamp & " and used to pend programs on PROG."
			End If


			If ButtonPressed = check_cash_box_btn Then CASH_checkbox = checked
			If ButtonPressed = check_grh_box_btn Then GRH_checkbox = checked
			If ButtonPressed = check_snap_box_btn Then SNAP_checkbox = checked
			If ButtonPressed = check_snap_box_btn Then snap_with_mfip = False
			If ButtonPressed = check_hc_box_btn Then HC_checkbox = checked
			If ButtonPressed = check_emer_box_btn Then EMER_checkbox = checked

			If ButtonPressed = check_cash_box_btn or ButtonPressed = check_grh_box_btn or ButtonPressed = check_snap_box_btn or ButtonPressed = check_hc_box_btn or ButtonPressed = check_emer_box_btn Then err_msg = "LOOP"

			If ButtonPressed = untrack_cash_btn Then CASH_checkbox = unchecked
			If ButtonPressed = untrack_grh_btn Then GRH_checkbox = unchecked
			If ButtonPressed = untrack_snap_btn Then SNAP_checkbox = unchecked
			If ButtonPressed = untrack_snap_btn Then snap_with_mfip = original_snap_with_mfip
			If ButtonPressed = untrack_hc_btn Then HC_checkbox = unchecked
			If ButtonPressed = untrack_emer_btn Then EMER_checkbox = unchecked

			If ButtonPressed = untrack_cash_btn or ButtonPressed = untrack_grh_btn or ButtonPressed = untrack_snap_btn or ButtonPressed = untrack_hc_btn or ButtonPressed = untrack_emer_btn Then err_msg = "LOOP"


			IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
	If REVW_CAF_Form = "MNbenefits" Then REVW_CAF_Form = "CAF (DHS-5223) from MNbenefits"
	If PROG_CAF_Form = "MNbenefits" Then PROG_CAF_Form = "CAF (DHS-5223) from MNbenefits"

	interview_required = True
	If option_to_process_with_no_interview = True and adult_cash_er_no_interview_checkbox = checked Then interview_required = False

	If CAF_datestamp = "" or (interview_required = True and interview_date = "") Then
	'This script is to support work after the interview and is not built to support the intervieww. Script will end if interview date is not found.
		end_early_mgs = "This script (NOTES - CAF) does not support details about an interview and should only be run once STAT panels are updated."
		end_early_mgs = end_early_mgs & vbCr & vbCr & "The script could not find details about the interview date. Update PROG or REVW with the correct interview date. Ensure all other STAT panels are updated and run NOTES - CAF again to document details about the information entered into STAT."
		end_early_mgs = end_early_mgs & vbCr & vbCr & "* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
		If CAF_datestamp = "" Then end_early_mgs = end_early_mgs & vbCr & vbCr & "FORM DATE HAS NOT BEEN ENTERED."
		If interview_required = True and interview_date = "" Then end_early_mgs = end_early_mgs & vbCr & vbCr & "INTERVIEW DATE HAS NOT BEEN ENTERED, AND THE INTERVIEW IS REQUIRED."
		end_early_mgs = end_early_mgs & vbCr & vbCr & "* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"
		Call script_end_procedure_with_error_report(end_early_mgs)
	End if

	If type_of_cash = "Family" Then
		adult_cash = FALSE
		family_cash = TRUE
	End If
	If type_of_cash = "Adult" Then
		adult_cash = TRUE
		family_cash = FALSE
	End If

	application_processing = False
	recert_processing = False
	If CASH_checkbox = checked and the_process_for_cash = "Application" Then application_processing = True
	If CASH_checkbox = checked and the_process_for_cash = "Recertification" Then recert_processing = True
	If GRH_checkbox = checked and the_process_for_grh = "Application" Then application_processing = True
	If GRH_checkbox = checked and the_process_for_grh = "Recertification" Then recert_processing = True
	If SNAP_checkbox = checked and the_process_for_snap = "Application" Then application_processing = True
	If SNAP_checkbox = checked and the_process_for_snap = "Recertification" Then recert_processing = True
	If EMER_checkbox = checked Then application_processing = True

	script_run_lowdown = script_run_lowdown & vbCr & vbCr & "DETAILS after PROGRAM Selection Dialog"
	If CASH_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH checkbox - CHECKED"
	If CASH_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "CASH checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "type_of_cash - " & type_of_cash
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_cash - " & the_process_for_cash
	script_run_lowdown = script_run_lowdown & vbCr & "cash_recert_mo - " & cash_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "cash_recert_yr - " & cash_recert_yr
	If GRH_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "GRH checkbox - CHECKED"
	If GRH_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "GRH checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_grh - " & the_process_for_grh
	script_run_lowdown = script_run_lowdown & vbCr & "grh_recert_mo - " & grh_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "grh_recert_yr - " & grh_recert_yr
	If SNAP_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP checkbox - CHECKED"
	If SNAP_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "SNAP checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "snap_with_mfip - " & snap_with_mfip
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_snap - " & the_process_for_snap
	script_run_lowdown = script_run_lowdown & vbCr & "snap_recert_mo - " & snap_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "snap_recert_yr - " & snap_recert_yr
	If HC_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "HC checkbox - CHECKED"
	If HC_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "HC checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_hc - " & the_process_for_hc
	script_run_lowdown = script_run_lowdown & vbCr & "hc_recert_mo - " & hc_recert_mo
	script_run_lowdown = script_run_lowdown & vbCr & "hc_recert_yr - " & hc_recert_yr
	If EMER_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER checkbox - CHECKED"
	If EMER_checkbox = unchecked Then script_run_lowdown = script_run_lowdown & vbCr & "EMER checkbox - UNCHECKED"
	script_run_lowdown = script_run_lowdown & vbCr & "type_of_emer - " & type_of_emer
	script_run_lowdown = script_run_lowdown & vbCr & "the_process_for_emer - " & the_process_for_emer

	script_run_lowdown = script_run_lowdown & vbCr & "REVW_CAF_Form - " & REVW_CAF_Form
	script_run_lowdown = script_run_lowdown & vbCr & "PROG_CAF_Form - " & PROG_CAF_Form

	script_run_lowdown = script_run_lowdown & vbCr & "application_processing - " & application_processing
	script_run_lowdown = script_run_lowdown & vbCr & "recert_processing - " & recert_processing

	script_run_lowdown = script_run_lowdown & vbCr & "Footer month: " & MAXIS_footer_month & "/" & MAXIS_footer_year

    Call back_to_SELF

    exp_det_case_note_found = False                         'defaulting these boolean variables to know if these notes are needed by this script run
    interview_completed_case_note_found = False
    verifications_requested_case_note_found = False
    caf_qualifying_questions_case_note_found = False

    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
    MAXIS_footer_year = right("00" & MAXIS_footer_year, 2)
    call check_for_MAXIS(False)	'checking for an active MAXIS session
    MAXIS_footer_month_confirmation	'function will check the MAXIS panel footer month/year vs. the footer month/year in the dialog, and will navigate to the dialog month/year if they do not match.

    'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    loop_start = timer
    Do
        call navigate_to_MAXIS_screen("STAT", "SUMM")
        EMReadScreen SUMM_check, 4, 2, 46
        Call back_to_SELF
        If timer - loop_start > 300 Then script_end_procedure("Can't get in to STAT. The script has attempted for 5 mintutes to get into STAT and iit appears to be stuck. The script timed out.")
    Loop until SUMM_check = "SUMM"

    'Creating a custom dialog for determining who the HH members are
    call HH_comp_dialog(HH_member_array)

    day_30_from_application = DateAdd("d", 30, CAF_datestamp)

    Call hest_standards(heat_AC_amt, electric_amt, phone_amt, CAF_datestamp)        'getting the correct amounts for HEST standards based on the application date.
    If the_process_for_snap = "Recertification" AND snap_recert_mo = "10" Then      'IF we are working a recertification CASE for SNAP for 10 - the recert month matters more than the app date. Pulling the correct HEST standards by footer month.
        oct_first_date = snap_recert_mo & "/1/" & snap_recert_yr
        oct_first_date = DateAdd("d", 0, oct_first_date)
        Call hest_standards(heat_AC_amt, electric_amt, phone_amt, oct_first_date)
    End If

    'THIS IS HANDLING SPECIFICALLY AROUND THE ALLOWANCE TO WAIVE INTERVIEWS FOR RENEWALS IN EFFECT STARTING FOR 04/21 REVW
    exp_screening_note_found = False

    MAXIS_case_number = trim(MAXIS_case_number)

    'HERE WE SEARCH CASE:NOTES
    'We are looking for notes that multiple scripts complete to keep from making duplicate notes
    look_for_expedited_determination_case_note = False
	If SNAP_checkbox = checked and the_process_for_snap = "Application" Then look_for_expedited_determination_case_note = True

    Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
    too_old_date = DateAdd("D", -1, CAF_datestamp)              'We don't need to read notes from before the CAF date

    note_row = 5
    Do
        EMReadScreen note_date, 8, note_row, 6                  'reading the note date

        EMReadScreen note_title, 55, note_row, 25               'reading the note header
        note_title = trim(note_title)

        'EXPEDITED DETERMINATION notes'
        If look_for_expedited_determination_case_note = True Then
            If left(note_title, 31) = "~ Received Application for SNAP" Then
                exp_screening_note_found = True
                EMWriteScreen "X", note_row, 3
                transmit

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
            End If
            If left(note_title, 47) = "Expedited Determination: SNAP appears expedited" Then                'reading a EXP CNote
                exp_det_case_note_found = TRUE
                snap_exp_yn = "Yes"
            End If
            If left(note_title, 55) = "Expedited Determination: SNAP does not appear expedited" Then        'Reading NOT EXP CNote
                exp_det_case_note_found = TRUE
                snap_exp_yn = "No"
            End If
            If left(note_title, 42) = "Expedited Determination: SNAP to be denied" Then                     'Reading DENY SNAP at EXP CNote
                exp_det_case_note_found = TRUE

                EMWriteScreen "X", note_row, 3  'Opens the note to read the denial date'
                transmit

                read_row = ""
                EMReadScreen find_denial_date_line, 22, 5, 3
                If find_denial_date_line = "* SNAP to be denied on" Then
                    read_row = 5
                Else
                    EMReadScreen find_denial_date_line, 22, 6, 3
                    If find_denial_date_line = "* SNAP to be denied on" Then read_row = 6
                End If
                If read_row <> "" Then
                    EMReadScreen note_denial_date, 10, row, 25
                    note_denial_date = replace(note_denial_date, "", ".")
                    note_denial_date = replace(note_denial_date, "", "S")
                    note_denial_date = replace(note_denial_date, "", "i")
                    note_denial_date = replace(note_denial_date, "", "n")
                    note_denial_date = replace(note_denial_date, "", "c")
                    note_denial_date = replace(note_denial_date, "", "e")
                    note_denial_date = trim(note_denial_date)
                    If IsDate(note_denial_date) = True Then snap_denial_date = note_denial_date
                End If

                PF3                             'closing the note
            End IF
        End If

        'INTERVIEW CNOTE
        If left(note_title, 24) = "~ Interview Completed on" Then
            interview_completed_case_note_found = True

            EMWriteScreen "X", note_row, 3                          'Opening the Interview Note to read some interview details
            transmit

            EMReadScreen in_correct_note, 24, 4, 3                  'making sure we are in the right note
            EMReadScreen note_list_header, 23, 4, 25

            If in_correct_note = "~ Interview Completed on" Then
                in_note_row = 5
                Do
                    EMReadScreen first_part_of_line, 12, in_note_row, 3                         'Reading the header portion
                    EMReadScreen whole_note_line, 77, in_note_row, 3                            'Reading all the line information
                    whole_note_line = trim(whole_note_line)
                    If first_part_of_line = "Completed wi" Then                                 'COMPLETED WITH header has person and type information
                        whole_note_line = replace(whole_note_line, "Completed with ", "")       'removes the header
                        position = Instr(whole_note_line, " via")                               'finds the dividing point in the content which is always the word 'via'
                        with_len = position + 4

                        interview_with = left(whole_note_line, position)                        'reading the person that did the interview - which is to the left of the dividing point
                        interview_type = right(whole_note_line, len(whole_note_line) - with_len)'the type is anything that is NOT the with
                        interview_with = trim(interview_with)
                        interview_type = trim(interview_type)
                    End If
                    If first_part_of_line = "Completed on" Then                                 'COMPLETED ON header has interview date
                        whole_note_line = replace(whole_note_line, "Completed on ", "")         'removes the header
                        position = Instr(whole_note_line, " at")                                'finds the dividing point

                        interview_date = left(whole_note_line, position)                        'interview date is to the left of the dividing point'
                        interview_date = trim(interview_date)
                    End If
                    in_note_row = in_note_row + 1
                    If interview_with <> "" AND interview_type <> "" AND interview_date <> "" Then Exit Do      'if we found all of it, we can be done
                Loop until trim(whole_note_line) = ""
                PF3         'leaving the note.

            Else
                If note_list_header <> "First line of Case note" Then PF3           'this backs us out of the note if we ended up in the wrong note.
            End If
        End If

        'VERIFICATIONS NOTES
        If left(note_title, 23) = "VERIFICATIONS REQUESTED" Then
            verifications_requested_case_note_found = True
            verifs_needed = "PREVIOUS NOTE EXISTS"

            EMWriteScreen "X", note_row, 3                          'Opening the VERIF note to read the verifications
            transmit

            EMReadScreen in_correct_note, 23, 4, 3                  'making sure we are in the right note
            EMReadScreen note_list_header, 23, 4, 25

            'Here we find the right row to start reading
            If in_correct_note = "VERIFICATIONS REQUESTED" Then                     'making sure we're in the right note
                in_note_row = 5
                Do
                    EMReadScreen whole_note_line, 77, in_note_row, 3                'reading the whole line of the note'
                    whole_note_line = trim(whole_note_line)

                    in_note_row = in_note_row + 1
                    If whole_note_line = "" then Exit Do
                Loop until whole_note_line = "List of all verifications requested:" 'This is the header within the note - the NEXT line starts the list of verifs

                If whole_note_line = "List of all verifications requested:" Then    'If we actually found the header.
                    verif_note_lines = ""                                           'defaulting a variable to save all the lines of the note
                    Do
                        EMReadScreen verif_line, 77, in_note_row, 3                 'reading the line of the note
                        verif_line = trim(verif_line)
                        If verif_line = "" then Exit Do                             'If they are blank - we stop'
                        verif_note_lines = verif_note_lines & "~|~" & verif_line    'Adding it to a string of all the lines

                        in_note_row = in_note_row + 1                               'next line'

                        EMReadScreen next_line, 77, in_note_row, 3                  'Looking to see if the next line is the divider line
                        next_line = trim(next_line)
                    Loop until next_line = "---"                                    'stop at the dividing line
                    'if there were lines saved
                    If verif_note_lines <> "" Then
                        verif_counter = 1                                           'setting a counter to find verifs that have been numbered
                        If left(verif_note_lines, 3) = "~|~" Then verif_note_lines = right(verif_note_lines, len(verif_note_lines) - 3)             'making an array of all of the lines
                        If InStr(verif_note_lines, "~|~") = 0 Then
                            verif_lines_array = Array(verif_note_lines)
                        Else
                            verif_lines_array = split(verif_note_lines, "~|~")
                        End If

                        verifs_to_add = ""                                          'blanking a string for adding all the lines together
                        For each line in verif_lines_array
                            counter_string = verif_counter & ". "                   'using the counter - which is a number to make a string that looks like what is in the note
                            If left(line, 2) = "- " OR left(line, 3) = counter_string Then                          'If the string starts with a dash or the counter
                                If left(line, 2) = "- " Then line = "; " & right(line, len(line) - 2)               'Removes the list delimiter and adds the editbox delimiter
                                If left(line, 3) = counter_string Then line = "; " & right(line, len(line) - 3)
                                verif_counter = verif_counter + 1                                                   'incrembting the counter
                            Else
                                line = " " & line                                                                   'adding a space to the sting so there is a space between words if we are at a 'same line'
                            End If

                            verifs_to_add = verifs_to_add & line                    'adding the verif information all together
                        Next
                        If left(verifs_to_add, 2) = "; " Then verifs_to_add = right(verifs_to_add, len(verifs_to_add) - 2)  'triming the string
                        If verifs_to_add <> "" Then verifs_needed = verifs_needed & " - Detail from NOTE: " & verifs_to_add 'adding the information to the variable used in this script
                    End If
                End If
                PF3         'leaving the note
            Else
                If note_list_header <> "First line of Case note" Then PF3           'this backs us out of the note if we ended up in the wrong note.
            End If
        End If

        'CAF QUALIFYING QUESTIONS NOTES
        If left(note_title, 47) = "CAF Qualifying Questions had an answer of 'YES'" Then
            caf_qualifying_questions_case_note_found = True

            EMWriteScreen "X", note_row, 3                          'Opening the CAF Qual Questions Note
            transmit

            EMReadScreen in_correct_note, 47, 4, 3                  'making sure we are in the right note
            EMReadScreen note_list_header, 23, 4, 25

            If in_correct_note = "CAF Qualifying Questions had an answer of 'YES'" Then
                in_note_row = 5
                Do
                    EMReadScreen whole_note_line, 77, in_note_row, 3                'Reading each line of the note
                    whole_note_line = trim(whole_note_line)

                    'Looks for the specific header for each of the qualifying questions
                    'If found it will default the droplist to Yes and pull the person information from the rest of the note line and saves it to the variable for the person
                    If left(whole_note_line, 42) = "* Fraud/DISQ for IPV (program violation): " Then
                        qual_question_one = "Yes"
                        qual_memb_one = replace(whole_note_line, "* Fraud/DISQ for IPV (program violation): ", "")
                    End If
                    If left(whole_note_line, 31) = "* SNAP in more than One State: " Then
                        qual_question_two = "Yes"
                        qual_memb_two = replace(whole_note_line, "* SNAP in more than One State: ", "")
                    End If
                    If left(whole_note_line, 17) = "* Fleeing Felon: " Then
                        qual_question_three = "Yes"
                        qual_memb_three = replace(whole_note_line, "* Fleeing Felon: ", "")
                    End If
                    If left(whole_note_line, 15) = "* Drug Felony: " Then
                        qual_question_four = "Yes"
                        qual_memb_four = replace(whole_note_line, "* Drug Felony: ", "")
                    End If
                    If left(whole_note_line, 30) = "* Parole/Probation Violation: " Then
                        qual_question_five = "Yes"
                        qual_memb_five = replace(whole_note_line, "* Parole/Probation Violation: ", "")
                    End If

                    in_note_row = in_note_row + 1
                    If whole_note_line = "" then Exit Do
                Loop until whole_note_line = "---"
                PF3
            Else
                If note_list_header <> "First line of Case note" Then PF3           'this backs us out of the note if we ended up in the wrong note.
            End If
        End If

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

        if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

        note_row = note_row + 1
        if note_row = 19 then
            note_row = 5
            PF8
            EMReadScreen check_for_last_page, 9, 24, 14
            If check_for_last_page = "LAST PAGE" Then Exit Do
        End If
        EMReadScreen next_note_date, 8, note_row, 6
        if next_note_date = "        " then Exit Do
    Loop until DateDiff("d", too_old_date, next_note_date) <= 0

    ' call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
    Call access_ADDR_panel("READ", notes_on_address, addr_line_one, addr_line_two, resi_street_full, city, state, zip, addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city_line, mail_state_line, mail_zip_line, addr_eff_date, addr_future_date, phone_number_one, phone_number_two, phone_number_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
    call read_SHEL_panel
    call update_shel_notes
    call read_HEST_panel
    ' call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)

    'Now it grabs the rest of the info, not dependent on which programs are selected.
    call autofill_editbox_from_MAXIS(HH_member_array, "ABPS", ABPS)
    call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", ACCI)
    call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", notes_on_acct)
    call autofill_editbox_from_MAXIS(HH_member_array, "ACUT", notes_on_acut)
    call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
    call autofill_editbox_from_MAXIS(HH_member_array, "BILS", BILS)
    call autofill_editbox_from_MAXIS(HH_member_array, "CASH", notes_on_cash)
    call autofill_editbox_from_MAXIS(HH_member_array, "CARS", notes_on_cars)
    call autofill_editbox_from_MAXIS(HH_member_array, "COEX", notes_on_coex)
    call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", notes_on_dcex)
    If cash_checkbox = checked Then call autofill_editbox_from_MAXIS(HH_member_array, "DIET", DIET)
    call autofill_editbox_from_MAXIS(HH_member_array, "DISA", DISA)
    call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
    call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
    call autofill_editbox_from_MAXIS(HH_member_array, "FMED", FMED)
    call autofill_editbox_from_MAXIS(HH_member_array, "IMIG", IMIG)
    call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
    'call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)

    job_count = 0
    call navigate_to_MAXIS_screen("STAT", "JOBS")
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check <> "0 Of 0" Then
        For each HH_member in HH_member_array
            EMWriteScreen HH_member, 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            EMReadScreen JOBS_total, 1, 2, 78
            If JOBS_total <> 0 then
                Do
                    ReDim Preserve ALL_JOBS_PANELS_ARRAY(budget_explain, job_count)
                    ALL_JOBS_PANELS_ARRAY(memb_numb, job_count) = HH_member
                    ALL_JOBS_PANELS_ARRAY(info_month, job_count) = MAXIS_footer_month & "/" & MAXIS_footer_year
                    call read_JOBS_panel

                    EMReadScreen JOBS_panel_current, 1, 2, 73
                    ALL_JOBS_PANELS_ARRAY(panel_instance, job_count) = "0" & JOBS_panel_current

                    If cint(JOBS_panel_current) < cint(JOBS_total) then transmit
                    job_count = job_count + 1
                Loop until cint(JOBS_panel_current) = cint(JOBS_total)
            End if
        Next
    End If

    If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then
        For each_memb = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
            Call Navigate_to_MAXIS_screen("CASE", "NOTE")

            too_old_date = DateAdd("D", -7, CAF_datestamp)
            ALL_JOBS_PANELS_ARRAY(EI_case_note, each_memb) = FALSE

            note_row = 5
            Do
                EMReadScreen note_date, 8, note_row, 6

                EMReadScreen note_title, 55, note_row, 25
                note_title = trim(note_title)

                If left(note_title, 14) = "INCOME DETAIL:" Then
                    member_reference = mid(note_title, 17, 2)
                    len_emp_name = len(ALL_JOBS_PANELS_ARRAY(employer_name, each_memb))
                    jobs_employer_name = mid(note_title, 29, len_emp_name)
                    jobs_employer_name = UCase(jobs_employer_name)

                    If member_reference = ALL_JOBS_PANELS_ARRAY(memb_numb, each_memb) AND jobs_employer_name = UCase(ALL_JOBS_PANELS_ARRAY(employer_name, each_memb)) Then
                        ALL_JOBS_PANELS_ARRAY(EI_case_note, each_memb) = TRUE
                    End If
                End If

                if note_date = "        " then Exit Do
                if ALL_JOBS_PANELS_ARRAY(EI_case_note, each_memb) = TRUE = TRUE then Exit Do

                note_row = note_row + 1
                if note_row = 19 then
                    'MsgBox "Next Page" & vbNewLine & "Note Date:" & note_date
                    note_row = 5
                    PF8
                    EMReadScreen check_for_last_page, 9, 24, 14
                    If check_for_last_page = "LAST PAGE" Then Exit Do
                End If
                EMReadScreen next_note_date, 8, note_row, 6
                if next_note_date = "        " then Exit Do
            Loop until DateDiff("d", too_old_date, next_note_date) <= 0
        Next
    End If

    busi_count = 0
    call navigate_to_MAXIS_screen("STAT", "BUSI")
    EMReadScreen panel_total_check, 6, 2, 73
    IF panel_total_check <> "0 Of 0" Then
        For each HH_member in HH_member_array
            EMWriteScreen HH_member, 20, 76
            EMWriteScreen "01", 20, 79
            transmit
            EMReadScreen BUSI_total, 1, 2, 78
            If BUSI_total <> 0 then

                Do
                    ReDim Preserve ALL_BUSI_PANELS_ARRAY(budget_explain, busi_count)
                    ALL_BUSI_PANELS_ARRAY(memb_numb, busi_count) = HH_member
                    ALL_BUSI_PANELS_ARRAY(info_month, busi_count) = MAXIS_footer_month & "/" & MAXIS_footer_year
                    call read_BUSI_panel

                    EMReadScreen BUSI_panel_current, 1, 2, 73
                    ALL_BUSI_PANELS_ARRAY(panel_instance, busi_count) = "0" & BUSI_panel_current

                    If cint(BUSI_panel_current) < cint(BUSI_total) then transmit
                    busi_count = busi_count + 1
                Loop until cint(BUSI_panel_current) = cint(BUSI_total)

            End if
        Next
    Else

    End If

    'FOR EACH JOB PANEL GO LOOK FOR A RECENT EI CASE NOTE'
    call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", INSA)
    call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", cit_id)
    call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", other_assets)
    call autofill_editbox_from_MAXIS(HH_member_array, "PBEN", case_changes)
    call autofill_editbox_from_MAXIS(HH_member_array, "PREG", PREG)
    call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
    call autofill_editbox_from_MAXIS(HH_member_array, "REST", notes_on_rest)
    call autofill_editbox_from_MAXIS(HH_member_array, "SCHL", SCHL)
    call autofill_editbox_from_MAXIS(HH_member_array, "SECU", other_assets)
    call autofill_editbox_from_MAXIS(HH_member_array, "STWK", notes_on_jobs)
    call read_TIME_panel
    call read_SANC_panel
    ' call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

    Call read_UNEA_panel

    call read_WREG_panel
    call update_wreg_and_abawd_notes
    'call autofill_editbox_from_MAXIS(HH_member_array, "WREG", notes_on_abawd)

    'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
    If cash_checkbox = checked then programs_applied_for = programs_applied_for & "CASH, "
    If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
    If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "emergency, "
    programs_applied_for = trim(programs_applied_for)
    if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

    'SHOULD DEFAULT TO TIKLING FOR APPLICATIONS THAT AREN'T RECERTS.
    If application_processing = True then TIKL_checkbox = checked

    Call generate_client_list(interview_memb_list, "Select or Type")
    Call generate_client_list(shel_memb_list, "Select")
    Call generate_client_list(verification_memb_list, "Select or Type Member")
    verification_memb_list = " "+chr(9)+verification_memb_list

    Call navigate_to_MAXIS_screen("STAT", "AREP")
    EMReadScreen version_numb, 1, 2, 73
    If version_numb = "1" Then
        EMReadScreen arep_name, 37, 4, 32
        arep_name = replace(arep_name, "_", "")
        interview_memb_list = interview_memb_list+chr(9)+"AREP - " & arep_name
    End If
End If
'This script is to support work after the interview and is not built to support the intervieww. Script will end if interview date is not found.
If interview_required = True and interview_date = "" Then
    end_early_mgs = "This script (NOTES - CAF) does not support details about an interview and should only be run once STAT panels are updated."
    end_early_mgs = end_early_mgs & vbCr & vbCr & "The script could not find details about the interview date. Update PROG or REVW with the correct interview date. Ensure all other STAT panels are updated and run NOTES - CAF again to document details about the information entered into STAT."
    Call script_end_procedure_with_error_report(end_early_mgs)
End if

prev_err_msg = ""
Do
    Do
        Do
            Do
                Do
                    Do
                        Do
                            Do
                                Do
                                    Do
                                        tab_button = False
                                        full_err_msg = ""
                                        err_array = ""
                                        If show_one = true Then
                                            dlg_len = 285
                                            For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                              If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
                                                  If ALL_MEMBERS_ARRAY(clt_age, the_member) > 17 OR ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then dlg_len = dlg_len + 20
                                              End If
                                            Next

                                            Dialog1 = ""
                                            ' BeginDialog Dialog1, 0, 0, 555, 385, "CAF Dialog 1 - Personal Information"
                                            BeginDialog Dialog1, 0, 0, 465, dlg_len, "CAF Dialog 1 - Personal Information"
                                              If interview_date <> "" Then Text 5, 10, 300, 10,  "* CAF datestamp:                               Interview date: " & interview_date
                                              If interview_date = "" Then Text 5, 10, 300, 10,  "* CAF datestamp: "
                                              If interview_required = False Then Text 5, 25, 300, 10, "No interview required for this " & CAF_form & " to be processed."
                                              If interview_required = True Then Text 5, 25, 300, 10, "Interview has been completed and documented previously."
                                              Text 5, 35, 300, 10, "Information about Case and Process Details:"

                                              EditBox 60, 5, 50, 15, CAF_datestamp
                                              EditBox 5, 45, 455, 15, case_details_and_notes_about_process

                                              Text 5, 65, 450, 10, "Member Name                         ID Type                              Detail                                                                                   Required"
                                              y_pos = 80
                                              For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                                If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
                                                    ' MsgBox "Name: " & ALL_MEMBERS_ARRAY(clt_name, the_member) & vbNewLine & "Age: " & ALL_MEMBERS_ARRAY(clt_age, the_member)
                                                    If ALL_MEMBERS_ARRAY(clt_age, the_member) > 17 OR ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then
                                                        If ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then ALL_MEMBERS_ARRAY(id_required, the_member) = checked
                                                        Text 5, y_pos, 85, 10, ALL_MEMBERS_ARRAY(clt_name, the_member)
                                                        ComboBox 100, y_pos - 5, 80, 15, "Type or Select"+chr(9)+"BC - Birth Certificate"+chr(9)+"RE - Religious Record"+chr(9)+"DL - Drivers License/ST ID"+chr(9)+"DV - Divorce Decree"+chr(9)+"AL - Alien Card"+chr(9)+"AD - Arrival//Depart"+chr(9)+"DR - Doctor Stmt"+chr(9)+"PV - Passport/Visa"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Verif Prvd", ALL_MEMBERS_ARRAY(clt_id_verif, the_member)
                                                        EditBox 185, y_pos - 5, 180, 15, ALL_MEMBERS_ARRAY(id_detail, the_member)
                                                        CheckBox 370, y_pos, 90, 10, "ID Verification Required", ALL_MEMBERS_ARRAY(id_required, the_member)
                                                        y_pos = y_pos + 20
                                                    End If
                                                End If
                                              Next
                                              Text 5, y_pos, 25, 10, "Citizen:"
                                              EditBox 35, y_pos -5, 425, 15, cit_id
                                              y_pos = y_pos + 20
                                              EditBox 35, y_pos - 5, 425, 15, IMIG
                                              y_pos = y_pos + 20
                                              EditBox 60, y_pos - 5, 120, 15, AREP
                                              EditBox 270, y_pos - 5, 190, 15, SCHL
                                              y_pos = y_pos + 20
                                              EditBox 60, y_pos - 5, 210, 15, DISA
                                              EditBox 310, y_pos - 5, 150, 15, FACI
                                              y_pos = y_pos + 20
                                              EditBox 35, y_pos - 5, 425, 15, PREG
                                              y_pos = y_pos + 20
                                              EditBox 35, y_pos - 5, 290, 15, ABPS
                                              If trim(ABPS) <> "" AND the_process_for_cash = "Application" Then
                                                Text 335, y_pos, 75, 10, "* Date CS Forms Sent:"
                                              Else
                                                Text 335, y_pos, 75,10, "Date CS Forms Sent:"
                                              End If
                                              EditBox 410, y_pos - 5, 35, 15, CS_forms_sent_date
                                              ButtonGroup ButtonPressed
                                                PushButton 445, y_pos - 5, 15, 15, "!", tips_and_tricks_cs_forms_button
                                              y_pos = y_pos + 20
                                              Text 5, y_pos, 30, 10, "Changes:"
                                              EditBox 40, y_pos - 5, 420, 15, case_changes
                                              y_pos = y_pos + 20 '210'
                                              EditBox 60, y_pos - 5, 385, 15, verifs_needed
                                              Text 10, y_pos + 50, 350, 10, "1 - Personal    |                    |                   |                   |                    |                   |                      |"
                                              ButtonGroup ButtonPressed
                                                PushButton 445, y_pos - 5, 15, 15, "!", tips_and_tricks_verifs_button
                                                PushButton 5, y_pos, 50, 10, "Verifs needed:", verif_button
                                                PushButton 60, y_pos + 50, 35, 10, "2 - JOBS", dlg_two_button
                                                PushButton 100, y_pos + 50, 35, 10, "3 - BUSI", dlg_three_button
                                                PushButton 140, y_pos + 50, 35, 10, "4 - CSES", dlg_four_button
                                                PushButton 180, y_pos + 50, 35, 10, "5 - UNEA", dlg_five_button
                                                PushButton 220, y_pos + 50, 35, 10, "6 - Other", dlg_six_button
                                                PushButton 260, y_pos + 50, 40, 10, "7 - Assets", dlg_seven_button
                                                PushButton 305, y_pos + 50, 50, 10, "8 - Interview", dlg_eight_button
                                                PushButton 370, y_pos + 45, 35, 15, "NEXT", go_to_next_page
                                                CancelButton 410, y_pos + 45, 50, 15
                                                PushButton 235, 15, 45, 10, "prev. panel", prev_panel_button
                                                PushButton 285, 15, 45, 10, "next panel", next_panel_button
                                                PushButton 345, 15, 45, 10, "prev. memb", prev_memb_button
                                                PushButton 395, 15, 45, 10, "next memb", next_memb_button
                                                PushButton 5, y_pos - 120, 20, 10, "IMIG:", IMIG_button
                                                PushButton 5, y_pos - 100, 25, 10, "AREP/", AREP_button
                                                PushButton 30, y_pos - 100, 25, 10, "ALTP:", ALTP_button
                                                PushButton 190, y_pos - 100, 25, 10, "SCHL/", SCHL_button
                                                PushButton 215, y_pos - 100, 25, 10, "STIN/", STIN_button
                                                PushButton 240, y_pos - 100, 25, 10, "STEC:", STEC_button
                                                PushButton 5, y_pos - 80, 25, 10, "DISA/", DISA_button
                                                PushButton 30, y_pos - 80, 25, 10, "PDED:", PDED_button
                                                PushButton 280, y_pos - 80, 25, 10, "FACI:", FACI_button
                                                PushButton 5, y_pos - 60, 25, 10, "PREG:", PREG_button
                                                PushButton 5, y_pos - 40, 25, 10, "ABPS:", ABPS_button
                                                PushButton 10, y_pos + 25, 20, 10, "DWP", ELIG_DWP_button
                                                PushButton 30, y_pos + 25, 15, 10, "FS", ELIG_FS_button
                                                PushButton 45, y_pos + 25, 15, 10, "GA", ELIG_GA_button
                                                PushButton 60, y_pos + 25, 15, 10, "HC", ELIG_HC_button
                                                PushButton 75, y_pos + 25, 20, 10, "MFIP", ELIG_MFIP_button
                                                PushButton 95, y_pos + 25, 20, 10, "MSA", ELIG_MSA_button
                                                PushButton 130, y_pos + 25, 25, 10, "ADDR", ADDR_button
                                                PushButton 155, y_pos + 25, 25, 10, "MEMB", MEMB_button
                                                PushButton 180, y_pos + 25, 25, 10, "MEMI", MEMI_button
                                                PushButton 205, y_pos + 25, 25, 10, "PROG", PROG_button
                                                PushButton 230, y_pos + 25, 25, 10, "REVW", REVW_button
                                                PushButton 255, y_pos + 25, 25, 10, "SANC", SANC_button
                                                PushButton 280, y_pos + 25, 25, 10, "TIME", TIME_button
                                                PushButton 305, y_pos + 25, 25, 10, "TYPE", TYPE_button
                                                If prev_err_msg <> "" Then PushButton 360, y_pos + 25, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                                OkButton 600, y_pos + 300, 50, 15
                                              GroupBox 5, y_pos + 15, 115, 25, "ELIG panels:"
                                              GroupBox 125, y_pos + 15, 210, 25, "other STAT panels:"
                                              GroupBox 230, 5, 215, 25, "STAT-based navigation"
                                              GroupBox 5, y_pos + 40, 355, 25, "Dialog Tabs"
                                            EndDialog

                                            Dialog Dialog1
                                            save_your_work
                                            cancel_confirmation
                                            MAXIS_dialog_navigation

                                            For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                                If ALL_MEMBERS_ARRAY(id_required, the_member) = checked AND ALL_MEMBERS_ARRAY(clt_id_verif, the_member) = "NO - No Verif Prvd" Then
                                                    verif_text = "Identity for Memb " & ALL_MEMBERS_ARRAY(memb_numb, the_member)
                                                    If InStr(verifs_needed, verif_text) = 0 Then verifs_needed = verifs_needed & "Identity for Memb " & ALL_MEMBERS_ARRAY(memb_numb, the_member) & ".; "
                                                End If
                                            Next

                                            verification_dialog

                                            If ButtonPressed = tips_and_tricks_interview_button Then tips_msg = MsgBox("*** Interview Detail ***" & vbNewLine & vbNewLine & "In order to actually process a CAF for all situations except one, an interview mst be completed. The CAF cannot be processed without an interview. This is why interview information is mandatory." & vbNewLine & vbNewLine &_
                                                                                                                       "An adult cash program ONLY at recertification is the only situation where the interview is not required. Any SNAP or application processing requires an interview." & vbNewLine & vbNewLine &_
                                                                                                                       "If an interview has not been completed, use either Client Contact to indicate the attempt to reach a client for an interview or Application Check to note information about a pending case.", vbInformation, "Tips and Tricks")
                                            If ButtonPressed = tips_and_tricks_cs_forms_button Then tips_msg = MsgBox("*** Date CS Forms Sent ***" & vbNewLine & vbNewLine & "For a Family Cash application and if there is information in the ABPS field, the script will require a date entered here." & vbNewLine & vbNewLine &_
                                                                                                                      "For family cash cases that are being denied enter 'N/A' to have the script bypass this field. Otherwise the date is required here." & vbNewLine & vbNewLine &_
                                                                                                                      "This field can also be used if the forms are given, instead of sent.", vbInformation, "Tips and Tricks")
                                            If ButtonPressed = tips_and_tricks_verifs_button Then tips_msg = MsgBox("*** Verifications Needed ***" & vbNewLine & vbNewLine & "This portion of the script has special functionality. Anytime this field is in a dialog, it is preceeded by a button instead of text." & vbNewLine & "** Press the button to open a special dialog to select verifications." & vbNewLine & vbNewLine &_
                                                                                                                    "Detail about this field/functionality:" & vbNewLine & " - The text '[Information here creates a SEPARATE CASE?NOTE]' can either be deleted or left in place. The script will ignore that phrase when entering a case note. The phrase must be exactly as is for the script to ignore." & vbNewLine &_
                                                                                                                    " - Use a '; ' - semi-colon followed by a space - to have the script go to the next line for the case note - great for formatting the case note." & vbNewLine & " - You can always type directly into the field by the button - you are not required to use the prepared checkboxes on other dialogs." & vbNewLine & vbNewLine &_
                                                                                                                    "VERIFICATIONS ARE ENTERED IN A SEPARATE CASE/NOTE. Do not list other case information in this field. Use 'Other Notes' or fields specific to the information to add.", vbInformation, "Tips and Tricks")
                                            ' If ButtonPressed = tips_and_tricks_interview_button Then ButtonPressed = dlg_one_button
                                            ' If ButtonPressed = tips_and_tricks_cs_forms_button Then ButtonPressed = dlg_one_button
                                            ' If ButtonPressed = tips_and_tricks_verifs_button Then ButtonPressed = dlg_one_button
                                            If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                            If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                                            Call assess_button_pressed
                                            If ButtonPressed = go_to_next_page Then pass_one = true
                                            If ButtonPressed = verif_button then
                                                pass_one = false
                                                show_one = true
                                            End If
                                            Dim Dialog1
                                        End If
                                    Loop Until pass_one = TRUE
                                    If show_two = true Then
                                        all_jobs = UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                        all_jobs = all_jobs + 1
                                        jobs_pages = all_jobs/3
                                        If jobs_pages <> Int(jobs_pages) Then jobs_pages = Int(jobs_pages) + 1

                                        each_job = 0
                                        loop_start = 0
                                        job_limit = 2
                                        Do
                                            last_job_reviewed = FALSE

                                            dlg_len = 85
                                            jobs_grp_len = 80
                                            length_factor = 80
                                            If snap_checkbox = checked Then length_factor = length_factor + 20
                                            If grh_checkbox = checked Then length_factor = length_factor + 20
                                            If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) = "" Then
                                                dlg_len = 100
                                            Else
                                                If UBound(ALL_JOBS_PANELS_ARRAY, 2) >= job_limit Then
                                                    dlg_len = 325
                                                    If snap_checkbox = checked Then dlg_len = dlg_len + 60
                                                    If grh_checkbox = checked Then dlg_len = dlg_len + 60
                                                    'jobs_grp_len = 315
                                                Else
                                                    dlg_len = length_factor * (UBound(ALL_JOBS_PANELS_ARRAY, 2) - loop_start + 1) + dlg_len
                                                    'jobs_grp_len = 100 * (UBound(ALL_JOBS_PANELS_ARRAY, 2) - loop_start + 1) + 15
                                                End If
                                            End If
                                            If snap_checkbox = checked Then jobs_grp_len = jobs_grp_len + 20
                                            If grh_checkbox = checked Then jobs_grp_len = jobs_grp_len + 20
                                            ' each_job = loop_start
                                            ' Do
                                            '     dlg_len = dlg_len + 100
                                            '     jobs_grp_len = jobs_grp_len + 100
                                            '     if each_job = job_limit Then Exit Do
                                            '     each_job = each_job + 1
                                            ' Loop until each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                            y_pos = 5
                                            'MsgBox dlg_len
                                            Dialog1 = ""
                                            ' BeginDialog Dialog1, 0, 0, 555, 385, "CAF Dialog 2 - JOBS Information"
                                            BeginDialog Dialog1, 0, 0, 705, dlg_len, "CAF Dialog 2 - JOBS Information"
                                              'GroupBox 5, 5, 595, jobs_grp_len, "Earned Income"
                                              If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) = "" Then
                                                y_pos = y_pos + 5
                                                Text 10, y_pos, 590, 10, "There are no JOBS panels found on this case. The script could not pull JOBS details for a case note."
                                                Text 10, y_pos + 10, 590, 10, " ** If this case has income from job source(s) it is best to add the JOBS panels before running this script. **"
                                                Text 10, y_pos + 30, 50, 10, "JOBS Details:"
                                                EditBox 55, y_pos + 25, 545, 15, notes_on_jobs
                                                y_pos = y_pos + 50
                                              Else
                                                  each_job = loop_start
                                                  Do
                                                      GroupBox 5, y_pos, 695, jobs_grp_len, "Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
                                                      Text 180, y_pos, 200, 10, "Verif: " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job)
                                                      CheckBox 365, y_pos, 220, 10, "Check here if this income is not verified and is only an estimate.", ALL_JOBS_PANELS_ARRAY(estimate_only, each_job)
                                                      y_pos = y_pos + 20
                                                      IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then
                                                        Text 15, y_pos, 690, 10, "Verification:                                                                                                                                                   EARNED INCOME BUDGETING CASE NOTE FOUND                                             to list of verifs needed."
                                                      Else
                                                        Text 15, y_pos, 690, 10, "Verification:                                                                                                                                                                                                                                                                                       to list of verifs needed."
                                                      End If
                                                      EditBox 65, y_pos - 5, 250, 15, ALL_JOBS_PANELS_ARRAY(verif_explain, each_job)
                                                      CheckBox 595, y_pos-10, 100, 10, "Check here to add this JOB", ALL_JOBS_PANELS_ARRAY(verif_checkbox, each_job)
                                                      y_pos = y_pos + 20
                                                      Text 15, y_pos, 600, 10, "Hourly Wage:                              Retro - Income:                              Hours:                                   Prosp - Income:                               Hours:                  Pay Freq:"
                                                      EditBox 65, y_pos - 5, 40, 15, ALL_JOBS_PANELS_ARRAY(hrly_wage, each_job)
                                                      EditBox 170, y_pos - 5, 45, 15, ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job)
                                                      EditBox 250, y_pos - 5, 20, 15, ALL_JOBS_PANELS_ARRAY(retro_hours, each_job)
                                                      EditBox 370, y_pos - 5, 45, 15, ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job)
                                                      EditBox 450, y_pos - 5, 20, 15, ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job)
                                                      ComboBox 520, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly"+chr(9)+ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job), ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job)
                                                      y_pos = y_pos + 20
                                                      If snap_checkbox = checked Then
                                                          Text 15, y_pos, 600, 10, "SNAP PIC:   * Pay Date Amount:                                                                          * Prospective Amount:                                               Calculated:"
                                                          EditBox 125, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)
                                                          ComboBox 185, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly"+chr(9)+ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job), ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job)
                                                          EditBox 340, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)
                                                          EditBox 470, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(pic_calc_date, each_job)
                                                          y_pos = y_pos + 20
                                                      End If
                                                      If grh_checkbox = checked Then
                                                          Text 15, y_pos, 35, 10, "GRH PIC:"
                                                          Text 65, y_pos, 60, 10, "Pay Date Amount: "
                                                          EditBox 125, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(grh_pay_day_income, each_job)
                                                          ComboBox 185, y_pos - 5, 60, 45, "Type or select"+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly"+chr(9)+chr(9)+ALL_JOBS_PANELS_ARRAY(grh_pay_freq, each_job), ALL_JOBS_PANELS_ARRAY(grh_pay_freq, each_job)
                                                          Text 265, y_pos, 70, 10, "Prospective Amount:"
                                                          EditBox 340, y_pos - 5, 60, 15, ALL_JOBS_PANELS_ARRAY(grh_prosp_income, each_job)
                                                          Text 420, y_pos, 40, 10, "Calculated:"
                                                          EditBox 470, y_pos - 5, 50, 15, ALL_JOBS_PANELS_ARRAY(grh_calc_date, each_job)
                                                          y_pos = y_pos + 20
                                                      End If
                                                      If ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = FALSE Then
                                                        Text 10, y_pos, 55, 10, "* Explain Budget:"
                                                      Else
                                                        Text 15, y_pos, 55, 10, "Explain Budget:"
                                                      End If
                                                      EditBox 70, y_pos - 5, 620, 15, ALL_JOBS_PANELS_ARRAY(budget_explain, each_job)
                                                      y_pos = y_pos + 25
                                                      if each_job = job_limit Then Exit Do
                                                      each_job = each_job + 1
                                                  Loop until each_job = UBound(ALL_JOBS_PANELS_ARRAY, 2) + 1
                                                  Text 10, y_pos + 5, 70, 40, "JOBS Details:                                              Other Earned Income:"
                                                  EditBox 65, y_pos, 620, 15, notes_on_jobs
                                                  Y_pos = y_pos + 20
                                                  If prev_err_msg <> "" Then
                                                    EditBox 85, y_pos, 510, 15, earned_income
                                                  Else
                                                    EditBox 85, y_pos, 615, 15, earned_income
                                                  End If
                                                  y_pos = y_pos + 25
                                              End If
                                              y_pos = y_pos + 5
                                              GroupBox 5, y_pos - 10, 355, 25, "Dialog Tabs"
                                              Text 10, y_pos, 300, 10, "                       |   2 - JOBS   |                   |                   |                    |                   |                      |"

                                              ButtonGroup ButtonPressed
                                                PushButton 685, y_pos - 50, 15, 15, "!", tips_and_tricks_jobs_button
                                                If prev_err_msg <> "" Then PushButton 600, y_pos - 30, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                                PushButton 10, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                                PushButton 100, y_pos, 35, 10, "3 - BUSI", dlg_three_button
                                                PushButton 140, y_pos, 35, 10, "4 - CSES", dlg_four_button
                                                PushButton 180, y_pos, 35, 10, "5 - UNEA", dlg_five_button
                                                PushButton 220, y_pos, 35, 10, "6 - Other", dlg_six_button
                                                PushButton 260, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                                PushButton 305, y_pos, 50, 10, "8 - Interview", dlg_eight_button

                                                If jobs_pages >= 2 Then
                                                    If jobs_pages = 2 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 440, y_pos, 15, 10, "2"
                                                        End If
                                                    ElseIf jobs_pages = 3 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                        End If
                                                    ElseIf jobs_pages = 4 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                        ElseIf loop_start = 9 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            Text 465, y_pos, 15, 10, "4"
                                                        End If
                                                    ElseIf jobs_pages = 5 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 9 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            Text 465, y_pos, 15, 10, "4"
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                        ElseIf loop_start = 12 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            Text 480, y_pos, 15, 10, "5"
                                                        End If
                                                    ElseIf jobs_pages = 6 Then
                                                        If loop_start = 0 Then
                                                            Text 420, y_pos, 15, 10, "1"
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 3 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            Text 435, y_pos, 15, 10, "2"
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 6 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            Text 450, y_pos, 15, 10, "3"
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 9 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            Text 465, y_pos, 15, 10, "4"
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 12 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            Text 480, y_pos, 15, 10, "5"
                                                            PushButton 495, y_pos, 15, 10, "6", jobs_page_six
                                                        ElseIf loop_start = 15 Then
                                                            PushButton 420, y_pos, 15, 10, "1", jobs_page_one
                                                            PushButton 435, y_pos, 15, 10, "2", jobs_page_two
                                                            PushButton 450, y_pos, 15, 10, "3", jobs_page_three
                                                            PushButton 465, y_pos, 15, 10, "4", jobs_page_four
                                                            PushButton 480, y_pos, 15, 10, "5", jobs_page_five
                                                            Text 495, y_pos, 15, 10, "6"
                                                        End If
                                                    End If
                                                End If

                                                PushButton 610, y_pos - 5, 35, 15, "NEXT", go_to_next_page
                                                CancelButton 650, y_pos - 5, 50, 15
                                                OkButton 750, 500, 50, 15
                                            EndDialog

                                            dialog Dialog1
                                            save_your_work
                                            cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                                            MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)

                                            If ButtonPressed = tips_and_tricks_jobs_button Then tips_msg = MsgBox("*** Entering JOBS Information ***" & vbNewLine & vbNewLine & "* If SNAP is checked, the SNAP specific information is ALWAYS required. We need more detail of earned income information in CASE/NOTE and these fields assist with that documentation." & vbNewLine & vbNewLine &_
                                                                                                                  "* The EXPLAIN BUDGET field is very important as it is where you can detail the conversation you had with the client about the income. This conversation is crucial to correct budgeting of JOBS income." & vbNewLine & vbNewLine &_
                                                                                                                  "* If you run the Earned Income Budgeting for a script prior to using the CAF script. The script will find the Earned Income CASE/NOTE and indicate it on this dialog. If that note is present you do NOT need to complete 'Explain Budget' or the SNAP Information as that has been well detailed by the Earned Income Script." & vbNewLine & vbNewLine &_
                                                                                                                  "* If you check the box at the top of the job information, indicating the information is only an estimate, additional detail in the 'Explain Budget' is not required. However, it is recommended to add additional detail if there was any conversation that occured or if there is specific detail that cannot be captured on JOBS." & vbNewLine & vbNewLine &_
                                                                                                                  "** WHAT TO DO IF A JOB HAS ENDED **" & vbNewLine & "This has come up a lot with all the required fields in the JOBS Dialog." & vbNewLine & vbNewLine &_
                                                                                                                  "* Income end date and STWK will be captured when the script gathers information. They will be listed in the fields in this dialog." & vbNewLine & vbNewLine &_
                                                                                                                  "* All the same fields are still mandatory. Since this JOBS panel exists in MAXIS, we need to address it in the case note. If ongoing income is 0, you can list 0 for the SNAP income. Explain budget can detail information about the job and what the changes are." & vbNewLine & vbNewLine &_
                                                                                                                  "* If this income is no longer budgeted - the panel can be removed. Review program specific information but typically once the job is out of the budget month and a STWK panel exists, the JOBS can be deleted. If the panel does not exist - then no detail would need to be entered about the job. (The panel must be deleted PRIOR to the script run.)" & vbNewLine & vbNewLine &_
                                                                                                                  "Generally, we have too little information about earned income in our CASE/NOTEs, this dialog guides you through adding sufficient detail about earned inocme and how it should be budgeted. The more information - the better, so use all applicable and available fields and explain IN FULL.", vbInformation, "Tips and Tricks")
                                            If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                            If ButtonPressed = tips_and_tricks_jobs_button Then ButtonPressed = dlg_two_button
                                            If ButtonPressed = -1 Then ButtonPressed = go_to_next_page
                                            If each_job >= UBound(ALL_JOBS_PANELS_ARRAY, 2) Then last_job_reviewed = TRUE

                                            Call assess_button_pressed
                                            If tab_button = TRUE Then last_job_reviewed = TRUE
                                            If ButtonPressed = go_to_next_page AND last_job_reviewed = TRUE Then pass_two = true

                                            job_limit = job_limit + 3
                                            loop_start = loop_start + 3
                                            If ButtonPressed = jobs_page_one Then
                                                loop_start = 0
                                                job_limit = 2
                                            ElseIf ButtonPressed = jobs_page_two Then
                                                loop_start = 3
                                                job_limit = 5
                                            ElseIf ButtonPressed = jobs_page_three Then
                                                loop_start = 6
                                                job_limit = 8
                                            ElseIf ButtonPressed = jobs_page_four Then
                                                loop_start = 9
                                                job_limit = 11
                                            ElseIf ButtonPressed = jobs_page_five Then
                                                loop_start = 12
                                                job_limit = 14
                                            ElseIf ButtonPressed = jobs_page_six Then
                                                loop_start = 15
                                                job_limit = 17
                                            End If
                                        Loop until last_job_reviewed = TRUE

                                        For each_job = o to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                            If ALL_JOBS_PANELS_ARRAY(verif_checkbox, each_job) = checked Then
                                                If ALL_JOBS_PANELS_ARRAY(verif_added, each_job) <> TRUE Then verifs_needed = verifs_needed & "Income for Memb " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " at " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & ".; "
                                                ALL_JOBS_PANELS_ARRAY(verif_added, each_job) = TRUE
                                            End If
                                        Next

                                    End If
                                Loop Until pass_two = true
                                If show_three = true Then
                                    all_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                    all_busi = all_busi + 1
                                    busi_pages = all_busi
                                    If busi_pages <> Int(busi_pages) Then busi_pages = Int(busi_pages) + 1

                                    each_busi = 0
                                    loop_start = 0
                                    last_busi_reviewed = FALSE
                                    busi_limit = 0
                                    Do
                                        dlg_len = 65
                                        busi_grp_len = 145
                                        length_factor = 140
                                        If snap_checkbox = checked Then length_factor = length_factor + 60
                                        If cash_checkbox = checked OR EMER_checkbox = checked Then length_factor = length_factor + 40
                                        'NEED HANDLING FOR IF NO JOBS'
                                        If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) = "" Then
                                            dlg_len = 80
                                        Else
                                            dlg_len = dlg_len + length_factor
                                            ' If UBound(ALL_busi_PANELS_ARRAY, 2) >= busi_limit Then
                                            '     dlg_len = dlg_len + 65
                                            '     If snap_checkbox = checked Then dlg_len = dlg_len + 60
                                            '     If cash_checkbox = checked OR EMER_checkbox = checked Then dlg_len = dlg_len + 60
                                            '     'busi_grp_len = 315
                                            ' Else
                                            '     dlg_len = length_factor * (UBound(ALL_BUSI_PANELS_ARRAY, 2) - loop_start + 1) + 65
                                            '     'busi_grp_len = 100 * (UBound(ALL_BUSI_PANELS_ARRAY, 2) - loop_start + 1) + 15
                                            ' End If
                                        End If
                                        If snap_checkbox = checked Then busi_grp_len = busi_grp_len + 60
                                        If cash_checkbox = checked OR EMER_checkbox = checked Then busi_grp_len = busi_grp_len + 40
                                        ' each_busi = loop_start
                                        ' Do
                                        '     dlg_len = dlg_len + 100
                                        '     busi_grp_len = busi_grp_len + 100
                                        '     if each_busi = busi_limit Then Exit Do
                                        '     each_busi = each_busi + 1
                                        ' Loop until each_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                        y_pos = 5

                                        Dialog1 = ""
                                        BeginDialog Dialog1, 0, 0, 546, dlg_len, "CAF Dialog 3 - BUSI"
                                          If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) = "" Then
                                            Text 10, y_pos, 535, 10, "There are no BUSI panels found on this case. The script could not pull BUSI details for a case note."
                                            Text 10, y_pos + 10, 535, 10, " ** If this case has income from self employment it is best to add the BUSI panels before running this script. **"
                                            Text 10, y_pos + 30, 50, 10, "BUSI Details:"
                                            EditBox 65, y_pos + 25, 475, 15, notes_on_busi
                                            y_pos = u_pos + 50
                                          Else
                                              each_busi = loop_start
                                              Do
                                                  GroupBox 5, y_pos, 535, busi_grp_len, "Member " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & "    Type: " & ALL_BUSI_PANELS_ARRAY(busi_type, each_busi)
                                                  CheckBox 290, y_pos, 220, 10, "Check here if this income is not verified and is only an estimate.", ALL_BUSI_PANELS_ARRAY(estimate_only, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 60, 10, "BUSI Description:"
                                                  EditBox 80, y_pos - 5, 445, 15, ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 55, 10, "BUSI Structure:"
                                                  ComboBox 75, y_pos - 5, 150, 45, "Select or Type"+chr(9)+"Sole Proprietor"+chr(9)+"Partnership"+chr(9)+"LLC"+chr(9)+"S Corp"+chr(9)+chr(9)+ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi), ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi)
                                                  Text 245, y_pos, 55, 10, "Ownership Share"
                                                  EditBox 305, y_pos - 5, 20, 15, ALL_BUSI_PANELS_ARRAY(share_num, each_busi)
                                                  Text 325, y_pos, 5, 10, "/"
                                                  EditBox 330, y_pos - 5, 20, 15, ALL_BUSI_PANELS_ARRAY(share_denom, each_busi)
                                                  Text 365, y_pos, 50, 10, "Partners in HH:"
                                                  EditBox 420, y_pos - 5, 105, 15, ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 90, 10, "* Self Employment Method:"
                                                  DropListBox 105, y_pos - 5, 120, 45, "Select One"+chr(9)+"50% Gross Inc"+chr(9)+"Tax Forms", ALL_BUSI_PANELS_ARRAY(calc_method, each_busi)
                                                  Text 240, y_pos, 45, 10, "Choice Date:"
                                                  EditBox 290, y_pos - 5, 50, 15, ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi)
                                                  CheckBox 350, y_pos, 185, 10, "Check here if SE Method was discussed with client", ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, each_busi)
                                                  y_pos = y_pos + 20
                                                  Text 15, y_pos, 200, 10, "Reported Hours:     Retro-                     Prosp-"
                                                  EditBox 100, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi)
                                                  EditBox 160, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi)
                                                  Text 205, y_pos, 300, 10, "Minimum Wage Hours:      Retro-                    Prosp-                    Income Start Date:"
                                                  EditBox 315, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi)
                                                  EditBox 375, y_pos - 5, 25, 15, ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi)
                                                  EditBox 470, y_pos - 5, 50, 15, ALL_BUSI_PANELS_ARRAY(start_date, each_busi)
                                                  y_pos = y_pos + 20
                                                  If SNAP_checkbox = checked Then
                                                      Text 15, y_pos, 200, 10, "SNAP:          Gross Income:      Retro-"
                                                      EditBox 140, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi)
                                                      Text 195, y_pos, 25, 10, "Prosp-"
                                                      EditBox 225, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi)
                                                      Text 295, y_pos, 100, 10, "Expenses:      Retro-"
                                                      EditBox 365, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_ret_snap, each_busi)
                                                      Text 420, y_pos, 25, 10, "Prosp-"
                                                      EditBox 450, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi)
                                                      y_pos = y_pos + 20
                                                      Text 50, y_pos, 85, 10, "* Expenses not allowed:"
                                                      EditBox 140, y_pos - 5, 355, 15, ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi)
                                                      y_pos = y_pos + 20
                                                      Text 115, y_pos, 25, 10, "Verif:"
                                                      ComboBox 140, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi), ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi)
                                                      Text 340, y_pos, 25, 10, "Verif:"
                                                      ComboBox 365, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi), ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi)
                                                      y_pos = y_pos + 20
                                                  End If
                                                  If cash_checkbox = checked OR EMER_checkbox = checked Then
                                                      Text 15, y_pos, 200, 10, "Cash/Emer:    Gross Income:      Retro-"
                                                      EditBox 140, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi)
                                                      Text 195, y_pos, 25, 10, "Prosp-"
                                                      EditBox 225, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi)
                                                      Text 295, y_pos, 100, 10, "Expenses:     Retro-"
                                                      EditBox 365, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_ret_cash, each_busi)
                                                      Text 420, y_pos, 25, 10, "Prosp-"
                                                      EditBox 450, y_pos - 5, 45, 15, ALL_BUSI_PANELS_ARRAY(expense_pro_cash, each_busi)
                                                      y_pos = y_pos + 20
                                                      Text 115, y_pos, 25, 10, "Verif:"
                                                      ComboBox 140, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi), ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi)
                                                      Text 340, y_pos, 25, 10, "Verif:"
                                                      ComboBox 365, y_pos - 5, 130, 45, "Select or Type"+chr(9)+"Income Tax Returns"+chr(9)+"Receipts of Sales/Purch"+chr(9)+"Busi Records/Ledger"+chr(9)+"Pend Out State Verif"+chr(9)+"Other Document"+chr(9)+"No Verif Provided"+chr(9)+"Delayed Verif"+chr(9)+"Blank"+chr(9)+ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi), ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi)
                                                      y_pos = y_pos + 20
                                                  End If
                                                  Text 15, y_pos, 65, 10, "Verification Detail:"
                                                  EditBox 80, y_pos - 5, 445, 15, ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi)
                                                  y_pos = y_pos + 15
                                                  CheckBox 80, y_pos, 400, 10, "Check here if verification about this Self Employment is requested.", ALL_BUSI_PANELS_ARRAY(verif_checkbox, each_busi)
                                                  y_pos = y_pos + 15
                                                  Text 15, y_pos, 60, 10, "* Explain Budget:"
                                                  EditBox 80, y_pos - 5, 445, 15, ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi)
                                                  y_pos = y_pos + 25
                                                  if each_busi = busi_limit Then Exit Do
                                                  each_busi = each_busi + 1
                                              Loop until each_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2) + 1
                                              Text 10, y_pos, 50, 10, "BUSI Details:"
                                              If prev_err_msg <> "" Then
                                                EditBox 60, y_pos - 5, 360, 15, notes_on_busi
                                              Else
                                                EditBox 60, y_pos - 5, 465, 15, notes_on_busi
                                              End If
                                              y_pos = y_pos + 20
                                          End If
                                          y_pos = y_pos + 10
                                          GroupBox 5, y_pos - 10, 355, 25, "Dialog Tabs"
                                          Text 10, y_pos, 300, 10, "                       |                    |   3 - BUSI   |                   |                    |                   |                      |"
                                          ButtonGroup ButtonPressed
                                            If prev_err_msg <> "" Then PushButton 425, y_pos - 35, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                            PushButton 10, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                            PushButton 60, y_pos, 35, 10, "2 - JOBS", dlg_two_button
                                            PushButton 140, y_pos, 35, 10, "4 - CSES", dlg_four_button
                                            PushButton 180, y_pos, 35, 10, "5 - UNEA", dlg_five_button
                                            PushButton 220, y_pos, 35, 10, "6 - Other", dlg_six_button
                                            PushButton 260, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                            PushButton 305, y_pos, 50, 10, "8 - Interview", dlg_eight_button
                                            If busi_pages >= 2 Then
                                                If busi_pages = 2 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                    ElseIf loop_start = 1 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                    End If
                                                ElseIf busi_pages = 3 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                    End If
                                                ElseIf busi_pages = 4 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                    ElseIf loop_start = 9 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        Text 410, y_pos, 15, 10, "4"
                                                    End If
                                                ElseIf busi_pages = 5 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 9 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        Text 410, y_pos, 15, 10, "4"
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                    ElseIf loop_start = 12 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        Text 425, y_pos, 15, 10, "5"
                                                    End If
                                                ElseIf busi_pages = 6 Then
                                                    If loop_start = 0 Then
                                                        Text 365, y_pos, 15, 10, "1"
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 3 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        Text 380, y_pos, 15, 10, "2"
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 6 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        Text 395, y_pos, 15, 10, "3"
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 9 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        Text 410, y_pos, 15, 10, "4"
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 12 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        Text 425, y_pos, 15, 10, "5"
                                                        PushButton 435, y_pos, 15, 10, "6", busi_page_six
                                                    ElseIf loop_start = 15 Then
                                                        PushButton 360, y_pos, 15, 10, "1", busi_page_one
                                                        PushButton 375, y_pos, 15, 10, "2", busi_page_two
                                                        PushButton 390, y_pos, 15, 10, "3", busi_page_three
                                                        PushButton 405, y_pos, 15, 10, "4", busi_page_four
                                                        PushButton 420, y_pos, 15, 10, "5", busi_page_five
                                                        Text 440, y_pos, 15, 10, "6"
                                                    End If
                                                End If
                                            End If
                                            PushButton 450, y_pos - 5, 35, 15, "NEXT", go_to_next_page
                                            CancelButton 490, y_pos - 5, 50, 15
                                            PushButton 525, 5, 15, 15, "!", tips_and_tricks_busi_button
                                            OkButton 600, 500, 50, 15
                                        EndDialog


                                        dialog Dialog1
                                        save_your_work
                                        cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                                        MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)

                                        If ButtonPressed = tips_and_tricks_busi_button Then tips_msg = MsgBox("*** Self_Employment ***" & vbNewLine & vbNewLine & "There is a policy update around Self Employment for SNAP that was went into effect 08/2019. If you are unfamiliar, this dialog will have elements that seem incorrect. Review the new policy on SIR and in the Policy Manuals." & vbNewLine & vbNewLine &_
                                                                                                              "* Business Description - This is not a field in MAXIS and can be used to further identify the self employment in CASE/NOTE. This can assist in the next worker understanding more about this case situation, making budgeting information clear, and make it easier to find documentation of this business in the case file." & vbNewLine & vbNewLine &_
                                                                                                              "* Business Structure, Ownership share, and Partners in Household - these fields also hep with idetifying budgeting and the correct focumentation required and on file for the businees. These fields are not required, but very helpful in a complete documentation." & vbNewLine & vbNewLine &_
                                                                                                              "* SNAP BUSI Budget - The new policy requires that we review TAX forms if that is the verification receivved to identify any allowed tax deductions that are not allowed as a part of SNAP budgeting. This field 'Expenses not Allowed' is required, though if all are allowed, simply use this field to indicate the review was done and all are allowed." & vbNewLine & vbNewLine &_
                                                                                                              "Checking the box that says 'Check here if verification about this Self Employment is requested' will add a line to the 'Verifs Needed' about self employment for this HH Member. Use this instead of typing to pulling up the verification dialog.", vbInformation, "Tips and Tricks")
                                        If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                        If ButtonPressed = tips_and_tricks_busi_button Then ButtonPressed = dlg_three_button
                                        If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                                        If each_busi >= UBound(ALL_BUSI_PANELS_ARRAY, 2) Then last_busi_reviewed = TRUE
                                        each_busi = loop_start
                                        Do
                                            'busi_err_msg'
                                            'IF THERE IS AN EI CASE NOTE - DON'T WORRY ABOUT MUCH ERR HANDLING
                                            if each_busi = busi_limit Then Exit Do
                                            each_busi = each_busi + 1
                                        Loop until each_busi = UBound(ALL_BUSI_PANELS_ARRAY, 2)

                                        Call assess_button_pressed
                                        If tab_button = TRUE Then last_busi_reviewed = TRUE
                                        If ButtonPressed = go_to_next_page AND last_busi_reviewed = TRUE Then pass_three = true

                                        busi_limit = busi_limit + 1
                                        loop_start = loop_start + 1

                                        If ButtonPressed = busi_page_one Then
                                            loop_start = 0
                                            job_limit = 0
                                        ElseIf ButtonPressed = busi_page_two Then
                                            loop_start = 1
                                            job_limit = 1
                                        ElseIf ButtonPressed = busi_page_three Then
                                            loop_start = 2
                                            job_limit = 2
                                        ElseIf ButtonPressed = busi_page_four Then
                                            loop_start = 3
                                            job_limit = 3
                                        ElseIf ButtonPressed = busi_page_five Then
                                            loop_start = 4
                                            job_limit = 4
                                        ElseIf ButtonPressed = busi_page_six Then
                                            loop_start = 5
                                            job_limit = 5
                                        End If

                                    Loop until last_busi_reviewed = TRUE

                                    For each_job = o to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                        If ALL_BUSI_PANELS_ARRAY(verif_checkbox, each_job) = checked Then
                                            If ALL_BUSI_PANELS_ARRAY(verif_added, each_job) <> TRUE Then verifs_needed = verifs_needed & "Self Employment Income for Memb " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_job) & ".; "
                                            ALL_BUSI_PANELS_ARRAY(verif_added, each_job) = TRUE
                                        End If
                                    Next

                                End If
                            Loop Until pass_three = true
                            If show_four = true Then
                                show_cses_detail = FALSE
                                group_len = 75
                                'If SNAP_checkbox = checked Then group_len = group_len + 40
                                group_wide = 465
                                If SNAP_checkbox = checked Then group_wide = 765
                                number_of_cs_members = 0
                                For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                    If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                                        ' dlg_four_len = dlg_four_len + 70
                                        'If SNAP_checkbox = checked Then dlg_four_len = dlg_four_len + 40
                                        show_cses_detail = TRUE
                                        number_of_cs_members = number_of_cs_members + 1
                                    End If
                                Next
                                cs_pages = number_of_cs_members/4
                                If cs_pages <> Int(cs_pages) Then cs_pages = Int(cs_pages) + 1
                                If show_cses_detail = FALSE Then dlg_four_len = 100
                                If SNAP_checkbox = checked AND show_cses_detail = TRUE Then
                                    dlg_wide = 775
                                Else
                                    dlg_wide = 480
                                End If

                                loop_start = 0
                                last_cs_reviewed = FALSE
                                cs_limit = 4

                                Do
                                    dlg_four_len = 85
                                    cs_counter = 0
                                    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                        If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                                            If cs_counter >= loop_start Then dlg_four_len = dlg_four_len + 70
                                            ' MsgBox "Counter - " & cs_counter & vbNewLine & "Limit - " & cs_limit & vbNewLine & "Loop start - " & loop_start & vbNewLine & "Dlg len - " & dlg_four_len
                                            cs_counter = cs_counter + 1
                                        End If
                                        If cs_counter = cs_limit Then Exit For
                                    Next
                                    If show_cses_detail = FALSE Then dlg_four_len = 100
                                    y_pos = 5
                                    ' MsgBox "Number of CS members - " & number_of_cs_members
                                    Dialog1 = ""
                                    BeginDialog Dialog1, 0, 0, dlg_wide, dlg_four_len, "Dialog 4 - CSES"
                                      If show_cses_detail = FALSE Then
                                          Text 10, y_pos, 445, 10, "There are no UNEA panels for Child Support (08, 36, 39) and the script could not pull child support detail information."
                                          Text 10, y_pos + 10, 445, 10, " ** If this case has income from child support it is best to add the UNEA panels before running this script. **"
                                          y_pos = y_pos + 30
                                      Else
                                          cs_counter = 0
                                          For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                              If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                                                  If cs_counter >= loop_start Then
                                                      GroupBox 5, y_pos, group_wide, group_len, "Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb)
                                                      y_pos = y_pos + 15
                                                      Text 10, y_pos, 260, 10, "Direct Child Support:       Amt/Mo: $                        Notes:"
                                                      EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                                                      If SNAP_checkbox = checked Then EditBox 195, y_pos - 5, 570, 15, UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)
                                                      If SNAP_checkbox = unchecked Then EditBox 195, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)
                                                      y_pos = y_pos + 20
                                                      If SNAP_checkbox = checked Then
                                                        Text 10, y_pos, 600, 10, "Disb Child Support(36):   Amt/Mo: $                        Notes:                                                                                                        Months to Average:                            Prosp Budg Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 200, 15, UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)
                                                        EditBox 465, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb)
                                                        EditBox 580, y_pos - 5, 185, 15, UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb)
                                                      Else
                                                        Text 10, y_pos, 250, 10, "Disb Child Support(36):   Amt/Mo: $                        Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)
                                                      End If
                                                      y_pos = y_pos + 20

                                                      If SNAP_checkbox = checked Then
                                                        Text 10, y_pos, 600, 10, "Disb CS Arrears(39):        Amt/Mo: $                        Notes:                                                                                                        Months to Average:                            Prosp Budg Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 200, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)
                                                        EditBox 465, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb)
                                                        EditBox 580, y_pos - 5, 185, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb)
                                                      Else
                                                        Text 10, y_pos, 250, 10, "Disb CS Arrears(39):        Amt/Mo: $                        Notes:"
                                                        EditBox 125, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                                                        EditBox 195, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)
                                                      End If
                                                      y_pos = y_pos + 20
                                                  End If
                                                  cs_counter = cs_counter + 1
                                              End If
                                              If cs_counter = cs_limit Then Exit For

                                          Next
                                          y_pos = y_pos + 10
                                      End If
                                      Text 10, y_pos, 60, 10, "Other CSES Detail:"

                                      If SNAP_checkbox = checked AND show_cses_detail = TRUE Then
                                          If prev_err_msg <> "" Then
                                            EditBox 75, y_pos - 5, 580, 15, notes_on_cses
                                            ButtonGroup ButtonPressed
                                              PushButton 660, y_pos - 5, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                          Else
                                            EditBox 75, y_pos - 5, 685, 15, notes_on_cses
                                          End If
                                          y_pos = y_pos + 20
                                          EditBox 60, y_pos - 5, 700, 15, verifs_needed
                                      Else
                                          If prev_err_msg <> "" Then
                                            EditBox 75, y_pos - 5, 290, 15, notes_on_cses
                                            ButtonGroup ButtonPressed
                                              PushButton 370, y_pos - 5, 100, 15, "Show Dialog Review Message", dlg_revw_button
                                          Else
                                            EditBox 75, y_pos - 5, 395, 15, notes_on_cses
                                          End If
                                          y_pos = y_pos + 20
                                          EditBox 60, y_pos - 5, 410, 15, verifs_needed
                                      End If

                                      y_pos = y_pos + 25
                                      GroupBox 10, y_pos - 10, 355, 25, "Dialog Tabs"
                                      Text 15, y_pos, 300, 10, "                       |                    |                  |   4 - CSES   |                    |                   |                      |"
                                      ButtonGroup ButtonPressed
                                        PushButton 5, y_pos - 25, 50, 10, "Verifs needed:", verif_button
                                        PushButton 15, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                        PushButton 65, y_pos, 35, 10, "2 - JOBS", dlg_two_button
                                        PushButton 105, y_pos, 35, 10, "3 - BUSI", dlg_three_button
                                        PushButton 185, y_pos, 35, 10, "5 - UNEA", dlg_five_button
                                        PushButton 225, y_pos, 35, 10, "6 - Other", dlg_six_button
                                        PushButton 265, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                        PushButton 310, y_pos, 50, 10, "8 - Interview", dlg_eight_button

                                        If cs_pages >= 2 Then
                                            If cs_pages = 2 Then
                                                If loop_start = 0 Then
                                                    Text 375, y_pos, 15, 10, "1"
                                                    PushButton 385, y_pos, 15, 10, "2", cs_page_two
                                                ElseIf loop_start = 4 Then
                                                    PushButton 370, y_pos, 15, 10, "1", cs_page_one
                                                    Text 390, y_pos, 15, 10, "2"
                                                End If
                                            ElseIf cs_pages = 3 Then
                                                If loop_start = 0 Then
                                                    Text 365, y_pos, 15, 10, "1"
                                                    PushButton 375, y_pos, 15, 10, "2", cs_page_two
                                                    PushButton 390, y_pos, 15, 10, "3", cs_page_three
                                                ElseIf loop_start = 4 Then
                                                    PushButton 360, y_pos, 15, 10, "1", cs_page_one
                                                    Text 380, y_pos, 15, 10, "2"
                                                    PushButton 390, y_pos, 15, 10, "3", cs_page_three
                                                ElseIf loop_start = 8 Then
                                                    PushButton 360, y_pos, 15, 10, "1", cs_page_one
                                                    PushButton 375, y_pos, 15, 10, "2", cs_page_two
                                                    Text 395, y_pos, 15, 10, "3"
                                                End If
                                            End If
                                        End If

                                        If SNAP_checkbox = checked AND show_cses_detail = TRUE Then
                                            PushButton 700, y_pos - 5, 25, 15, "NEXT", go_to_next_page
                                            CancelButton 730, y_pos - 5, 30, 15
                                        Else
                                            PushButton 410, y_pos - 5, 25, 15, "NEXT", go_to_next_page
                                            CancelButton 440, y_pos - 5, 30, 15
                                        End If
                                        OkButton 700, 700, 50, 15
                                    EndDialog

                                    Dialog Dialog1
                                    save_your_work
                                    cancel_confirmation
                                    verification_dialog
                                    'MsgBox ButtonPressed
                                    If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                                    If ButtonPressed = -1 Then ButtonPressed = go_to_next_page
                                    If ButtonPressed = verif_button Then ButtonPressed = dlg_four_button
                                    If cs_counter >= number_of_cs_members Then last_cs_reviewed = TRUE

                                    Call assess_button_pressed
                                    If tab_button = TRUE Then last_cs_reviewed = TRUE
                                    If ButtonPressed = go_to_next_page AND last_cs_reviewed = TRUE Then pass_four = true

                                    cs_limit = cs_limit + 4
                                    loop_start = loop_start + 4
                                    If ButtonPressed = cs_page_one Then
                                        loop_start = 0
                                        cs_limit = 4
                                    ElseIf ButtonPressed = cs_page_two Then
                                        loop_start = 4
                                        cs_limit = 8
                                    ElseIf ButtonPressed = cs_page_three Then
                                        loop_start = 9
                                        cs_limit = 12
                                    End If
                                Loop until last_cs_reviewed = TRUE
                            End If
                        Loop Until pass_four = true
                        If show_five = true Then
                            dlg_five_len = 190
                            ssa_group_len = 30
                            uc_group_len = 30
                            unea_income_found = FALSE
                            For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                                    dlg_five_len = dlg_five_len + 70
                                    uc_group_len = uc_group_len + 80
                                    UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) & ""
                                    unea_income_found = TRUE
                                End If
                                If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
                                    dlg_five_len = dlg_five_len + 40
                                    ssa_group_len = ssa_group_len + 40
                                    unea_income_found = TRUE
                                End If
                            Next
                            If trim(notes_on_VA_income) <> "" Then unea_income_found = TRUE
                            If trim(notes_on_WC_income) <> "" Then unea_income_found = TRUE
                            If trim(notes_on_other_UNEA) <> "" Then unea_income_found = TRUE
                            If unea_income_found = FALSE Then dlg_five_len = dlg_five_len + 20

                            y_pos = 5
                            Dialog1 = ""
                            BeginDialog Dialog1, 0, 0, 466, dlg_five_len, "Dialog 5 - UNEA"
                              If unea_income_found = FALSE Then
                                  Text 10, y_pos, 445, 10, "There are no UNEA panels found and the script could not pull detail about SSA/WC/VA/UC or other UNEA income."
                                  Text 10, y_pos + 10, 445, 10, " ** If this case has income from SSI, RSDI, or Unemployment it is best to add the UNEA panels before running this script. **"
                                  y_pos = y_pos + 25
                              End If
                              GroupBox 5, y_pos, 455, ssa_group_len, "SSA Income"
                              y_pos = y_pos + 15
                              For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                  If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
                                      Text 15, y_pos, 40, 10, "Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb)
                                      Text 60, y_pos, 55, 10, "RSDI: Amount: $"
                                      EditBox 120, y_pos - 5, 30, 15, UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)
                                      Text 155, y_pos, 30, 10, "* Notes:"
                                      EditBox 185, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)
                                      y_pos = y_pos + 20
                                      Text 60, y_pos, 55, 10, "SSI: Amount: $"
                                      EditBox 120, y_pos - 5, 30, 15, UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)
                                      Text 155, y_pos, 30, 10, "* Notes:"
                                      EditBox 185, y_pos - 5, 270, 15, UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)
                                      y_pos = y_pos + 20
                                  End If
                              Next
                              Text 10, y_pos, 65, 10, "Other SSA Income:"
                              EditBox 80, y_pos - 5, 375, 15, notes_on_ssa_income
                              y_pos = y_pos + 25
                              Text 5, y_pos, 40, 10, "VA Income:"
                              EditBox 45, y_pos - 5, 415, 15, notes_on_VA_income
                              y_pos = y_pos + 20
                              Text 5, y_pos, 55, 10, "Worker's Comp:"
                              EditBox 60, y_pos - 5, 400, 15, notes_on_WC_income
                              y_pos = y_pos + 15
                              GroupBox 5, y_pos, 455, uc_group_len, "Unemployment Income"
                              y_pos = y_pos + 15
                              For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                  If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                                      UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ""
                                      UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ""

                                      Text 15, y_pos, 40, 10, "Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb)
                                      Text 65, y_pos, 120, 10, "Unemployment Start Date: " & UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb)
                                      Text 195, y_pos, 95, 10, "* Budgeted Weekly Amount:"
                                      EditBox 290, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)
                                      Text 345, y_pos, 70, 10, "UC Acct Bal:"
                                      EditBox 395, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)
                                      y_pos = y_pos + 20
                                      Text 30, y_pos, 50, 10, "Weekly Gross:"
                                      EditBox 85, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)
                                      Text 130, y_pos, 70, 10, "Allowed Deductions:"
                                      EditBox 200, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)
                                      Text 245, y_pos, 75, 10, "Excluded Deductions:"
                                      EditBox 320, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)
                                      Text 375, y_pos - 5, 80, 15, "Enter a TIKL date to check if UC has ended:"
                                      y_pos = y_pos + 20
                                      Text 30, y_pos, 50, 10, "Retro Income:"
                                      EditBox 80, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb)
                                      Text 130, y_pos, 50, 10, "Prosp Income:"
                                      EditBox 185, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb)
                                      If SNAP_checkbox = checked Then Text 250, y_pos, 65, 10, "* SNAP Prosp Amt: $"
                                      If SNAP_checkbox = unchecked Then Text 250, y_pos, 65, 10, "SNAP Prosp Amt: $"
                                      EditBox 315, y_pos - 5, 40, 15, UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)
                                      ButtonGroup ButtonPressed
                                        PushButton 365, y_pos, 35, 10, "Calc", calc_button
                                      EditBox 405, y_pos - 5, 50, 15, UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)
                                      y_pos = y_pos + 20
                                      Text 30, y_pos, 25, 10, "Notes:"
                                      EditBox 60, y_pos - 5, 395, 15, UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb)
                                      y_pos = y_pos + 20
                                  End If
                              Next
                              Text 15, y_pos, 60, 10, "Other UC Income:"
                              EditBox 75, y_pos - 5, 380, 15, other_uc_income_notes
                              y_pos = y_pos + 25
                              Text 10, y_pos, 45, 10, "Other UNEA:"
                              If prev_err_msg <> "" Then
                                EditBox 55, y_pos - 5, 305, 15, notes_on_other_UNEA
                                ButtonGroup ButtonPressed
                                  PushButton 365, y_pos - 5, 100, 15, "Show Dialog Review Message", dlg_revw_button
                              Else
                                EditBox 55, y_pos - 5, 405, 15, notes_on_other_UNEA
                              End If
                              y_pos = y_pos + 20
                              ButtonGroup ButtonPressed
                                PushButton 5, y_pos, 50, 10, "Verifs needed:", verif_button
                              EditBox 60, y_pos - 5, 400, 15, verifs_needed
                              y_pos = y_pos + 25
                              GroupBox 10, y_pos - 10, 355, 25, "Dialog Tabs"
                              Text 15, y_pos, 300, 10, "                       |                    |                   |                   |  5 - UNEA   |                   |                      |"
                              ButtonGroup ButtonPressed
                                PushButton 15, y_pos, 45, 10, "1 - Personal", dlg_one_button
                                PushButton 65, y_pos, 35, 10, "2 - JOBS", dlg_two_button
                                PushButton 105, y_pos, 35, 10, "3 - BUSI", dlg_three_button
                                PushButton 145, y_pos, 35, 10, "4 - CSES", dlg_four_button
                                PushButton 225, y_pos, 35, 10, "6 - Other", dlg_six_button
                                PushButton 265, y_pos, 40, 10, "7 - Assets", dlg_seven_button
                                PushButton 310, y_pos, 50, 10, "8 - Interview", dlg_eight_button
                                PushButton 370, y_pos - 5, 35, 15, "NEXT", go_to_next_page
                                CancelButton 410, y_pos - 5, 50, 15
                                OkButton 600, 500, 50, 15
                            EndDialog

                            Dialog Dialog1
                            save_your_work
                            cancel_confirmation
                            MAXIS_dialog_navigation
                            verification_dialog

                            If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                            If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                            If ButtonPressed = calc_button Then
                                For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                                    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                                        If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)) = TRUE Then
                                            If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) = TRUE Then
                                                weeks_of_UC_benefits = Int(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)/UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb))
                                                'MsgBox weeks_of_UC_benefits
                                                UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) = DateAdd("ww", weeks_of_UC_benefits, date)
                                            Else
                                                MsgBox "The script cannot calculate the potential date of UC account balance depletion for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " without the UC account balance and UC Weekly Gross income. Enter these amounts as numbers and the script will enter a date for the TIKL into the dialog. The TIKL date can also be entered or changed manually."
                                            End If
                                        End If
                                    End If
                                Next
                                ButtonPressed = dlg_five_button
                            End If
                            If ButtonPressed = verif_button then ButtonPressed = dlg_five_button

                            Call assess_button_pressed
                            If ButtonPressed = go_to_next_page Then pass_five = true
                        End If
                    Loop Until pass_five = true
                    If show_six = true Then
                        If left(total_shelter_amount, 1) <> "$" Then total_shelter_amount = "$" & total_shelter_amount
                        combined_electric_and_phone_amt = electric_amt + phone_amt
                        heat_ac_detail = "AC/Heat - Full $" & heat_AC_amt
                        electric_phone_detail = "Electric and Phone - $" & combined_electric_and_phone_amt
                        electric_detail = "Electric ONLY - $" & electric_amt
                        phone_detail = "Phone ONLY - $" & phone_amt
                        hest_droplist = "Select ALLOWED HEST"+chr(9)+heat_ac_detail+chr(9)+electric_phone_detail+chr(9)+electric_detail+chr(9)+phone_detail+chr(9)+"NONE - $0"

                        Dialog1 = ""
                        BeginDialog Dialog1, 0, 0, 556, 290, "CAF Dialog 6 - WREG, Expenses, Address"
                          EditBox 45, 50, 500, 15, notes_on_wreg
                          ButtonGroup ButtonPressed
                            PushButton 440, 30, 105, 15, "Update ABAWD and WREG", abawd_button
                            PushButton 235, 85, 50, 15, "Update SHEL", update_shel_button
                          DropListBox 45, 140, 100, 45, hest_droplist, hest_information
                          EditBox 180, 140, 110, 15, notes_on_acut
                          EditBox 45, 160, 245, 15, notes_on_coex
                          EditBox 45, 180, 245, 15, notes_on_dcex
                          EditBox 45, 200, 245, 15, notes_on_other_deduction
                          EditBox 45, 220, 245, 15, expense_notes
                          CheckBox 320, 85, 125, 10, "Check here to confirm the address.", address_confirmation_checkbox
                          DropListBox 345, 150, 85, 45, county_list, addr_county
                          DropListBox 480, 150, 30, 45, "No"+chr(9)+"Yes", homeless_yn
                          DropListBox 335, 170, 95, 45, "SF - Shelter Form"+chr(9)+"CO - Coltrl Stmt"+chr(9)+"LE - Lease/Rent Doc"+chr(9)+"MO - Mortgage Papers"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"CD - Contrct for Deed"+chr(9)+"UT - Utility Stmt"+chr(9)+"DL - Driver Lic/State ID"+chr(9)+"OT - Other Document"+chr(9)+"NO - No Ver Prvd"+chr(9)+"? - Delayed"+chr(9)+"Blank", addr_verif
                          DropListBox 480, 170, 30, 45, "No"+chr(9)+"Yes", reservation_yn
                          DropListBox 375, 190, 165, 45, "  "+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
                          EditBox 315, 220, 230, 15, notes_on_address
                          EditBox 60, 245, 490, 15, verifs_needed
                          GroupBox 5, 5, 545, 65, "WREG and ABAWD Information"
                          Text 15, 15, 55, 10, "ABAWD Details:"
                          Text 75, 15, 470, 10, notes_on_abawd
                          Text 15, 25, 400, 10, notes_on_abawd_two
                          Text 15, 35, 400, 10, notes_on_abawd_three
                          GroupBox 5, 75, 290, 165, "Expenses and Deductions"
                          Text 15, 90, 50, 10, "Total Shelter:"
                          Text 70, 90, 155, 10, total_shelter_amount
                          Text 10, 105, 285, 10, shelter_details
                          Text 10, 115, 285, 10, shelter_details_two
                          Text 10, 125, 285, 10, shelter_details_three
                          Text 20, 205, 20, 10, "Other:"
                          Text 20, 225, 25, 10, "Notes:"
                          GroupBox 305, 75, 245, 165, "Address"
                          Text 350, 100, 175, 10, addr_line_one
                          If addr_line_two = "" Then
                            Text 350, 115, 175, 10, city & ", " & state & " " & zip
                          Else
                            Text 350, 115, 175, 10, addr_line_two
                            Text 350, 130, 175, 10, city & ", " & state & " " & zip
                          End If
                          Text 315, 155, 25, 10, "County:"
                          Text 440, 155, 35, 10, "Homeless:"
                          Text 315, 175, 20, 10, "Verif:"
                          Text 435, 175, 45, 10, "Reservation:"
                          Text 315, 195, 55, 10, "* Living Situation:"
                          Text 315, 210, 75, 10, "Notes on address:"
                          GroupBox 105, 265, 355, 25, "Dialog Tabs"
                          Text 110, 275, 300, 10, "                       |                    |                   |                    |                    |  6 - Other   |                      |"
                          ButtonGroup ButtonPressed
                            PushButton 5, 250, 50, 10, "Verifs needed:", verif_button
                            If prev_err_msg <> "" Then PushButton 5, 270, 100, 15, "Show Dialog Review Message", dlg_revw_button
                            PushButton 110, 275, 45, 10, "1 - Personal", dlg_one_button
                            PushButton 160, 275, 35, 10, "2 - JOBS", dlg_two_button
                            PushButton 200, 275, 35, 10, "3 - BUSI", dlg_three_button
                            PushButton 240, 275, 35, 10, "4 - CSES", dlg_four_button
                            PushButton 280, 275, 35, 10, "5 - UNEA", dlg_five_button
                            PushButton 360, 275, 40, 10, "7 - Assets", dlg_seven_button
                            PushButton 405, 275, 50, 10, "8 - Interview", dlg_eight_button
                            PushButton 460, 270, 35, 15, "NEXT", go_to_next_page
                            CancelButton 500, 270, 50, 15
                            If SNAP_checkbox = checked Then PushButton 10, 55, 30, 10, "* WREG", wreg_button
                            If SNAP_checkbox = unchecked Then PushButton 10, 55, 25, 10, "WREG", wreg_button
                            PushButton 315, 100, 25, 10, "ADDR", addr_button
                            PushButton 15, 145, 25, 10, "HEST", hest_button
                            PushButton 150, 145, 25, 10, "ACUT", acut_button
                            PushButton 15, 165, 25, 10, "COEX", coex_button
                            PushButton 15, 185, 25, 10, "DCEX", dcex_button
                            OkButton 600, 500, 50, 15
                        EndDialog

                        Dialog Dialog1			'Displays the second dialog
                        save_your_work
                        cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
                        MAXIS_dialog_navigation			'Navigates around MAXIS using a custom function (works with the prev/next buttons and all the navigation buttons)
                        verification_dialog

                        If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                        If ButtonPressed = -1 Then ButtonPressed = go_to_next_page

                        If ButtonPressed = abawd_button Then
                            Do
                                abawd_err_msg = ""

                                notes_on_wreg = ""
                                notes_on_abawd = ""
                                notes_on_abawd_two = ""
                                notes_on_abawd_three = ""
                                dlg_len = 40
                                For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                  If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked AND ALL_MEMBERS_ARRAY(wreg_exists, each_member) = TRUE Then
                                    dlg_len = dlg_len + 95
                                  End If
                                Next
                                y_pos = 10
                                Dialog1 = ""
                                BeginDialog Dialog1, 0, 0, 551, dlg_len, "ABAWD Detail"
                                  For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                    If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked AND ALL_MEMBERS_ARRAY(wreg_exists, each_member) = TRUE Then
                                      GroupBox 5, y_pos, 540, 95, "Member " & ALL_MEMBERS_ARRAY(memb_numb, each_member) & " - " & ALL_MEMBERS_ARRAY(clt_name, each_member)
                                      y_pos = y_pos + 20
                                      Text 15, y_pos, 70, 10, "FSET WREG Status:"
                                      DropListBox 90, y_pos - 5, 130, 45, " "+chr(9)+"03  Unfit for Employment"+chr(9)+"04  Responsible for Care of Another"+chr(9)+"05  Age 60+"+chr(9)+"06  Under Age 16"+chr(9)+"07  Age 16-17, live w/ parent"+chr(9)+"08  Care of Child <6"+chr(9)+"09  Employed 30+ hrs/wk"+chr(9)+"10  Matching Grant"+chr(9)+"11  Unemployment Insurance"+chr(9)+"12  Enrolled in School/Training"+chr(9)+"13  CD Program"+chr(9)+"14  Receiving MFIP"+chr(9)+"20  Pend/Receiving DWP"+chr(9)+"15  Age 16-17 not live w/ Parent"+chr(9)+"16  50-59 Years Old"+chr(9)+"21  Care child < 18"+chr(9)+"17  Receiving RCA or GA"+chr(9)+"30  FSET Participant"+chr(9)+"02  Fail FSET Coop"+chr(9)+"33  Non-coop being referred"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(clt_wreg_status, each_member)
                                      Text 230, y_pos, 55, 10, "ABAWD Status:"
                                      DropListBox 285, y_pos - 5, 110, 45, " "+chr(9)+"01  WREG Exempt"+chr(9)+"02  Under Age 18"+chr(9)+"03  Age 50+"+chr(9)+"04  Caregiver of Minor Child"+chr(9)+"05  Pregnant"+chr(9)+"06  Employed 20+ hrs/wk"+chr(9)+"07  Work Experience"+chr(9)+"08  Other E and T"+chr(9)+"09  Waivered Area"+chr(9)+"10  ABAWD Counted"+chr(9)+"11  Second Set"+chr(9)+"12  RCA or GA Participant"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(clt_abawd_status, each_member)
                                      CheckBox 405, y_pos - 5, 130, 10, "Check here if this person is the PWE", ALL_MEMBERS_ARRAY(pwe_checkbox, each_member)
                                      y_pos = y_pos + 20
                                      Text 15, y_pos, 145, 10, "Number of ABAWD months used in past 36:"
                                      EditBox 160, y_pos - 5, 25, 15, ALL_MEMBERS_ARRAY(numb_abawd_used, each_member)
                                      Text 200, y_pos, 95, 10, "List all ABAWD months used:"
                                      EditBox 300, y_pos - 5, 135, 15, ALL_MEMBERS_ARRAY(list_abawd_mo, each_member)
                                      y_pos = y_pos + 20
                                      Text 15, y_pos, 135, 10, "If used, list the first month of Second Set:"
                                      EditBox 155, y_pos - 5, 40, 15, ALL_MEMBERS_ARRAY(first_second_set, each_member)
                                      Text 205, y_pos, 130, 10, "If NOT Eligible for Second Set, Explain:"
                                      EditBox 335, y_pos - 5, 200, 15, ALL_MEMBERS_ARRAY(explain_no_second, each_member)
                                      y_pos = y_pos + 20
                                      'Text 15, y_pos, 115, 10, "Number of BANKED months used:"
                                      'EditBox 130, y_pos - 5, 25, 15, ALL_MEMBERS_ARRAY(numb_banked_mo, each_member)
                                      Text 15, y_pos, 45, 10, "Other Notes:"
                                      EditBox 60, y_pos - 5, 475, 15, ALL_MEMBERS_ARRAY(clt_abawd_notes, each_member)

                                      y_pos = y_pos + 15
                                    End If
                                  Next
                                  y_pos = y_pos + 10
                                  ButtonGroup ButtonPressed
                                    PushButton 455, y_pos, 90, 15, "Return to Main Dialog", return_button
                                    OkButton 600, 500, 50, 15
                                EndDialog

                                Dialog Dialog1
                                save_your_work

                                If ButtonPressed = -1 Then ButtonPressed = return_button
                                If ButtonPressed = 0 Then ButtonPressed = return_button

                                call update_wreg_and_abawd_notes
                                If ButtonPressed = return_button Then ButtonPressed = dlg_six_button

                            Loop until abawd_err_msg = ""
                        End If

                        If ButtonPressed = update_shel_button Then
                            shel_client = ""
                            For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                If ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE Then
                                    shel_client = each_member
                                    Exit For
                                End If
                            Next
                            If shel_client <> "" Then clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, shel_client)
                            'ADD an IF here to determine the right HH member or if one is not yet selected AND preselect the one that has a SHEL'
                            Do
                                shel_err_msg = ""

                                If clt_SHEL_is_for = "Select" Then
                                    dlg_len = 30
                                Else
                                    dlg_len = 250
                                    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                                        If clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, each_member) Then
                                            shel_client = each_member
                                            ALL_MEMBERS_ARRAY(shel_exists, each_member) = TRUE
                                        End If
                                    Next
                                End If
                                if shel_client = "" Then
                                    shel_client = 0
                                    clt_shel_is_for = ALL_MEMBERS_ARRAY(full_clt, shel_client)
                                End If

                                If ALL_MEMBERS_ARRAY(shel_subsudized, shel_client) = "" Then ALL_MEMBERS_ARRAY(shel_subsudized, shel_client) = "No"
                                If ALL_MEMBERS_ARRAY(shel_shared, shel_client) = "" Then ALL_MEMBERS_ARRAY(shel_shared, shel_client) = "No"
                                shel_verif_needed_checkbox = unchecked
                                If manual_total_shelter = "" Then manual_total_shelter = total_shelter_amount & ""
                                If manual_amount_used = FALSE Then manual_total_shelter = total_shelter_amount & ""
                                start_total_shel = manual_total_shelter

                                Dialog1 = ""
                                BeginDialog Dialog1, 0, 0, 340, dlg_len, "SHEL Detail Dialog"
                                  DropListBox 60, 10, 125, 45, shel_memb_list, clt_SHEL_is_for
                                  Text 5, 15, 55, 10, "SHEL for Memb"
                                  ButtonGroup ButtonPressed
                                    PushButton 190, 10, 40, 10, "Load", load_button
                                  Text 235, 10, 55, 10, "Total Shelter:"
                                  EditBox 290, 5, 40, 15, manual_total_shelter
                                  If clt_shel_is_for <> "Select" Then
                                      'ALL_MEMBERS_ARRAY
                                      ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) & ""
                                      ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) & ""

                                      DropListBox 85, 30, 30, 45, "Yes"+chr(9)+"No", ALL_MEMBERS_ARRAY(shel_subsudized, shel_client)
                                      DropListBox 175, 30, 30, 45, "Yes"+chr(9)+"No", ALL_MEMBERS_ARRAY(shel_shared, shel_client)
                                      EditBox 45, 60, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client)
                                      DropListBox 85, 60, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_rent_verif, shel_client)
                                      EditBox 195, 60, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client)
                                      DropListBox 235, 60, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_rent_verif, shel_client)
                                      EditBox 45, 80, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client)
                                      DropListBox 85, 80, 100, 45, "Select one"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_lot_verif, shel_client)
                                      EditBox 195, 80, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client)
                                      DropListBox 235, 80, 100, 45, "Select one"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_lot_verif, shel_client)
                                      EditBox 45, 100, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client)
                                      DropListBox 85, 100, 100, 45, "Select one"+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_mortgage_verif, shel_client)
                                      EditBox 195, 100, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client)
                                      DropListBox 235, 100, 100, 45, "Select one"+chr(9)+"MO - Mort Pmt Book"+chr(9)+"CD - Ctrct For Deed"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_mortgage_verif, shel_client)
                                      EditBox 45, 120, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client)
                                      DropListBox 85, 120, 100, 45, "Select one"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_ins_verif, shel_client)
                                      EditBox 195, 120, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client)
                                      DropListBox 235, 120, 100, 45, "Select one"+chr(9)+"BI - Billing Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_ins_verif, shel_client)
                                      EditBox 45, 140, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client)
                                      DropListBox 85, 140, 100, 45, "Select one"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_tax_verif, shel_client)
                                      EditBox 195, 140, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client)
                                      DropListBox 235, 140, 100, 45, "Select one"+chr(9)+"TX - Prop Tax Stmt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_tax_verif, shel_client)
                                      EditBox 45, 160, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client)
                                      DropListBox 85, 160, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_room_verif, shel_client)
                                      EditBox 195, 160, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client)
                                      DropListBox 235, 160, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_room_verif, shel_client)
                                      EditBox 45, 180, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client)
                                      DropListBox 85, 180, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_garage_verif, shel_client)
                                      EditBox 195, 180, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client)
                                      DropListBox 235, 180, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"RE - Rent Receipt"+chr(9)+"OT - Other Doc"+chr(9)+"NC - Change - Neg Impact"+chr(9)+"PC - Change - Pos Impact"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_garage_verif, shel_client)
                                      EditBox 45, 200, 35, 15, ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client)
                                      DropListBox 85, 200, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_retro_subsidy_verif, shel_client)
                                      EditBox 195, 200, 35, 15, ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client)
                                      DropListBox 235, 200, 100, 45, "Select one"+chr(9)+"SF - Shelter Form"+chr(9)+"LE - Lease"+chr(9)+"OT - Other Doc"+chr(9)+"NO - No Verif"+chr(9)+"? - Delayed Verif"+chr(9)+"Blank", ALL_MEMBERS_ARRAY(shel_prosp_subsidy_verif, shel_client)
                                      CheckBox 45, 220, 150, 10, "Check here if verification is requested.", ALL_MEMBERS_ARRAY(shel_verif_checkbox, shel_client)
                                      CheckBox 45, 235, 185, 10, "Check here if this verification is NOT MANDATORY.", not_mand_checkbox
                                      ButtonGroup ButtonPressed
                                        PushButton 245, 230, 90, 15, "Return to Main Dialog", return_button
                                        OkButton 600, 500, 50, 15
                                      Text 15, 35, 60, 10, "HUD Subsidized:"
                                      Text 140, 35, 30, 10, "Shared:"
                                      Text 45, 50, 50, 10, "Retrospective"
                                      Text 195, 50, 50, 10, "Prospective"
                                      Text 20, 65, 20, 10, "Rent:"
                                      Text 10, 85, 30, 10, "Lot Rent:"
                                      Text 5, 105, 35, 10, "Mortgage:"
                                      Text 5, 125, 35, 10, "Insurance:"
                                      Text 15, 145, 25, 10, "Taxes:"
                                      Text 15, 165, 25, 10, "Room:"
                                      Text 10, 185, 30, 10, "Garage:"
                                      Text 10, 205, 30, 10, "Subsidy:"
                                  End If
                                EndDialog

                                dialog Dialog1
                                save_your_work

                                If IsNumeric(manual_total_shelter) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Total Shelter costs must be a number."
                                If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Lot Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Lot Rent Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Morgage Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective ortgage Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Insurance Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Insurance Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Tax Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Tax Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Room Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Room Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Garage Expense."
                                If ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Garage Expense."
                                If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Retro Subsidy Amount."
                                If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) <> "" AND IsNumeric(ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client)) = FALSE Then shel_err_msg = shel_err_msg & vbNewLine & "* Enter a valid amount for Prospective Subsidy Amount."

                                If ButtonPressed = load_button Then shel_err_msg = "LOOP" & shel_err_msg

                                If left(shel_err_msg, 4) <> "LOOP" AND shel_err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & shel_err_msg

                                If ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_rent_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_rent_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_lot_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_lot_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_mortgage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_mortgage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_ins_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_ins_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_tax_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_tax_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_room_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_room_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_garage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_garage_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_retro_subsidy_amt, shel_client) * 1
                                If ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) <> "" Then ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) = ALL_MEMBERS_ARRAY(shel_prosp_subsidy_amt, shel_client) * 1

                                call update_shel_notes
                                If ALL_MEMBERS_ARRAY(shel_verif_checkbox, shel_client) = checked Then
                                    If ALL_MEMBERS_ARRAY(shel_verif_added, shel_client) <> TRUE Then
                                        verifs_needed = verifs_needed & "Shelter costs for Memb " & ALL_MEMBERS_ARRAY(full_clt, shel_client) & ". "
                                        If not_mand_checkbox = checked Then verifs_needed = verifs_needed & " THIS VERIFICATION IS NOT MANDATORY."
                                        verifs_needed = verifs_needed & "; "
                                    End If
                                    ALL_MEMBERS_ARRAY(shel_verif_added, shel_client) = TRUE
                                End If

                                If ButtonPressed = -1 Then ButtonPressed = return_button
                                If ButtonPressed = 0 Then ButtonPressed = return_button

                                If ButtonPressed = return_button Then ButtonPressed = dlg_six_button
                                If manual_total_shelter <> start_total_shel Then
                                    manual_amount_used = TRUE
                                    total_shelter_amount = manual_total_shelter
                                End If
                                If manual_amount_used = TRUE Then total_shelter_amount = manual_total_shelter
                                total_shelter_amount = total_shelter_amount * 1
                            Loop until shel_err_msg = ""
                        End If
                        If ButtonPressed = verif_button then ButtonPressed = dlg_six_button

                        Call assess_button_pressed
                        If ButtonPressed = go_to_next_page Then pass_six = true
                    End If
                Loop Until pass_six = true
                If show_seven = true Then
                    app_month_assets = app_month_assets & ""

                    Dialog1 = ""
                    BeginDialog Dialog1, 0, 0, 561, 340, "CAF Dialog 7 - Asset and Miscellaneous Info"
                      EditBox 435, 20, 115, 15, app_month_assets
                      EditBox 45, 40, 395, 15, notes_on_acct
                      EditBox 475, 40, 75, 15, notes_on_cash
                      CheckBox 45, 60, 350, 10, "Check here to confirm NO account panels and all income was reviewed for direct deposit payments.", confirm_no_account_panel_checkbox
                      EditBox 45, 80, 235, 15, notes_on_cars
                      EditBox 315, 80, 235, 15, notes_on_rest
                      EditBox 115, 100, 435, 15, notes_on_other_assets
                      EditBox 40, 130, 275, 15, MEDI
                      EditBox 360, 130, 195, 15, DIET
                      EditBox 40, 150, 515, 15, FMED
                      EditBox 40, 170, 515, 15, DISQ
                      EditBox 40, 205, 510, 15, notes_on_time
                      EditBox 60, 225, 490, 15, notes_on_sanction
                      EditBox 50, 245, 500, 15, EMPS
                      ButtonGroup ButtonPressed
                        PushButton 25, 265, 15, 15, "!", tips_and_tricks_emps_button
                      EditBox 60, 290, 495, 15, verifs_needed
                      GroupBox 105, 310, 355, 25, "Dialog Tabs"
                      Text 110, 320, 300, 10, "                       |                    |                   |                  |                    |                    |   7 - Assets    |"
                      ButtonGroup ButtonPressed
                        PushButton 15, 45, 25, 10, "ACCT", acct_button
                        PushButton 445, 45, 25, 10, "CASH", cash_button
                        PushButton 10, 135, 25, 10, "MEDI:", MEDI_button
                        PushButton 325, 135, 25, 10, "DIET:", DIET_button
                        PushButton 10, 155, 25, 10, "FMED:", FMED_button
                        PushButton 15, 85, 25, 10, "CARS", cars_button
                        If InStr(shelter_details, "Mortgage") <> 0 Then PushButton 285, 85, 25, 10, "* REST", rest_button
                        If InStr(shelter_details, "Mortgage") = 0 Then PushButton 285, 85, 25, 10, "REST", rest_button
                        PushButton 15, 105, 25, 10, "SECU", secu_button
                        PushButton 40, 105, 25, 10, "TRAN", tran_button
                        PushButton 65, 105, 45, 10, "other assets", other_asset_button
                        PushButton 10, 175, 25, 10, "DISQ:", disq_button
                        If family_cash = TRUE Then PushButton 15, 250, 30, 10, "* EMPS:", emps_button
                        If family_cash = FALSE Then PushButton 20, 250, 25, 10, "EMPS:", emps_button
                        PushButton 5, 295, 50, 10, "Verifs needed:", verif_button
                        If prev_err_msg <> "" Then PushButton 450, 265, 100, 15, "Show Dialog Review Message", dlg_revw_button
                        PushButton 110, 320, 45, 10, "1 - Personal", dlg_one_button
                        PushButton 160, 320, 35, 10, "2 - JOBS", dlg_two_button
                        PushButton 200, 320, 35, 10, "3 - BUSI", dlg_three_button
                        PushButton 240, 320, 35, 10, "4 - CSES", dlg_four_button
                        PushButton 280, 320, 35, 10, "5 - UNEA", dlg_five_button
                        PushButton 320, 320, 35, 10, "6 - Other", dlg_six_button
                        PushButton 405, 320, 50, 10, "8 - Interview", dlg_eight_button
                        PushButton 465, 315, 35, 15, "NEXT", go_to_next_page
                        CancelButton 505, 315, 50, 15
                        OkButton 600, 500, 50, 15
                      GroupBox 10, 10, 545, 115, "Assets"
                      If the_process_for_snap = "Application" Then
                        Text 310, 25, 110, 10, "* Total Liquid Assets in App Month:"
                      Else
                        Text 310, 25, 110, 10, "Total Liquid Assets in App Month:"
                      End If
                      GroupBox 10, 190, 545, 95, "MFIP/DWP"
                      If family_cash = TRUE Then
                          Text 15, 210, 25, 10, "* Time:"
                          Text 15, 230, 35, 10, "* Sanction:"
                      Else
                          Text 20, 210, 20, 10, "Time:"
                          Text 20, 230, 30, 10, "Sanction:"
                      End If
                    EndDialog

                    Dialog Dialog1
                    save_your_work
                    cancel_confirmation
                    MAXIS_dialog_navigation
                    verification_dialog

                    If ButtonPressed = tips_and_tricks_emps_button Then tips_msg = MsgBox("*** TIME, Sanction, and EMPS ***" & vbNewLine & "Why are these now required?" & vbNewLine & vbNewLine &_
                                                                                          "Information about TIME (TANF months used), SANC (Details about MFIP Sanctions), and EMPS (MFIP Employment Services) are now required for any case that is Family Cash when running the CAF. These elements are paramount to the MFIP program and should be addressed at least once per year. Review of these peices of a case can go here." & vbNewLine & vbNewLine &_
                                                                                          "What if it is a new case?" & vbNewLine & "* This is a great place to indicate that there is no history of time or sanctions used, that the client reports no benefits in another state, or that you are waiting on detail from another state. This is also a good place to identify EMPS requirement was explained or DWP overview scheduled/completed." & vbNewLine & vbNewLine &_
                                                                                          "This is a relative caregiver case, why is it needed here?" & vbNewLine & "* Since these function differently for these cases, you may not be detailing time used. Detailing that it is specifically NOT being used is extremely helpful to the new HSR or reviewer that works on this case. Add detail about how these typically mandatory elements do NOT apply in this case." & vbNewLine & vbNewLine &_
                                                                                          "The script will try to autofill this information but additional detail is helpful as always.", vbInformation, "Tips and Tricks")

                    If ButtonPressed = dlg_revw_button THen Call display_errors(prev_err_msg, FALSE)
                    If ButtonPressed = tips_and_tricks_emps_button Then ButtonPressed = dlg_seven_button
                    If ButtonPressed = -1 Then ButtonPressed = go_to_next_page
                    If ButtonPressed = verif_button then ButtonPressed = dlg_seven_button

                    Call assess_button_pressed
                    If ButtonPressed = go_to_next_page Then pass_seven = true

                    If IsNumeric(app_month_assets) = TRUE Then app_month_assets = app_month_assets * 1
                End If
            Loop Until pass_seven = true
            If show_eight = true Then
                If the_process_for_snap = "Application" AND exp_det_case_note_found = FALSE Then
                    If full_determination_done = False Then
                        full_determination_done = True
                        If first_time_to_exp_det = True Then
                            determined_income = 0
                            For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                                If ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job) <> "" Then
                                    If IsNumeric(ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)) = True Then determined_income = determined_income + ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)
                                ElseIf ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job) <> "" Then
                                    If IsNumeric(ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)) = True Then
                                        If ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job) = "Monthly" Then determined_income = determined_income + ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)
                                        If ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job) = "Semi-Monthly" Then determined_income = determined_income + ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)*2
                                        If ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job) = "Biweekly" Then determined_income = determined_income + ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)*2.15
                                        If ALL_JOBS_PANELS_ARRAY(main_pay_freq, each_job) = "Weekly" Then determined_income = determined_income + ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)*4.3
                                    End If
                                End If
                            Next

                            For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                                If IsNumeric(ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi)) = True Then determined_income = determined_income + ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi)-ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi)
                            Next

                            determined_income = determined_income + determined_unea_income

                            determined_assets = app_month_assets

                            total_shelter_amount = replace(total_shelter_amount, "$", "")
                            total_shelter_amount = total_shelter_amount * 1
                            determined_shel = total_shelter_amount

                            If hest_information <> "Select ALLOWED HEST" Then
                                hest_array = split(hest_information, "-")
                                determined_utilities = hest_array(1)
                                determined_utilities = replace(determined_utilities, "$", "")
                                determined_utilities = replace(determined_utilities, "Full", "")
                                determined_utilities = trim(determined_utilities)
                                If IsNumeric(determined_utilities) = False then determined_utilities = ""
                            End If
                            first_time_to_exp_det = False
                        End If
                        Call run_expedited_determination_script_functionality(xfs_screening, caf_one_income, caf_one_assets, caf_one_rent, caf_one_utilities, determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, approval_date, CAF_datestamp, interview_date, applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, snap_denial_date, snap_denial_explain, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, postponed_verifs_yn, list_postponed_verifs, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, action_due_to_out_of_state_benefits, case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn, delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
                    End If
                End If
            End If
            If show_eight = true Then

                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 500, 370, "CAF Dialog 8 - Interview Info"
                  EditBox 60, 10, 20, 15, next_er_month
                  EditBox 85, 10, 20, 15, next_er_year
                  ComboBox 330, 10, 165, 15, "Select or Type"+chr(9)+"incomplete"+chr(9)+"approved"+chr(9)+CAF_status, CAF_status
                  EditBox 60, 30, 435, 15, actions_taken
                  ' DropListBox 135, 60, 30, 45, "?"+chr(9)+"Yes"+chr(9)+"No", snap_exp_yn
                  ' ButtonGroup ButtonPressed
                    ' PushButton 165, 60, 15, 15, "!", tips_and_tricks_xfs_button
                  ' EditBox 270, 60, 40, 15, app_month_income '210'
                  ' EditBox 350, 60, 40, 15, app_month_assets '290'
                  ' EditBox 445, 60, 40, 15, app_month_expenses '385'
                  ' EditBox 90, 80, 35, 15, exp_snap_approval_date
                  ' EditBox 195, 80, 295, 15, exp_snap_delays
                  ' EditBox 90, 100, 35, 15, snap_denial_date
                  ' EditBox 195, 100, 295, 15, snap_denial_explain
                  CheckBox 20, 155, 80, 10, "Application signed?", application_signed_checkbox
                  CheckBox 20, 170, 50, 10, "eDRS sent?", eDRS_sent_checkbox
                  CheckBox 20, 185, 65, 10, "Updated MMIS?", updated_MMIS_checkbox
                  CheckBox 20, 200, 95, 10, "Workforce referral made?", WF1_checkbox
                  CheckBox 125, 155, 85, 10, "Sent forms to AREP?", Sent_arep_checkbox
                  CheckBox 125, 170, 80, 10, "Intake packet given?", intake_packet_checkbox
                  CheckBox 125, 185, 70, 10, "IAAs/OMB given?", IAA_checkbox
                  CheckBox 220, 155, 115, 10, "Informed client of recert period?", recert_period_checkbox
                  CheckBox 220, 170, 130, 10, "Rights and Responsibilities explained?", R_R_checkbox
                  CheckBox 220, 185, 150, 10, "Client Requests to participate with E and T", E_and_T_checkbox
                  CheckBox 220, 200, 125, 10, "Eligibility Requirements Explained?", elig_req_explained_checkbox
                  CheckBox 220, 215, 160, 10, "Benefits and Payment Information Explained?", benefit_payment_explained_checkbox
                  EditBox 55, 240, 440, 15, other_notes
                  EditBox 60, 260, 435, 15, verifs_needed
                  CheckBox 15, 295, 240, 10, "Check here to update PND2 to show client delay (pending cases only).", client_delay_checkbox
                  CheckBox 15, 310, 200, 10, "Check here to create a TIKL to deny at the 30 day mark.", TIKL_checkbox
                  CheckBox 15, 325, 265, 10, "Check here to send a TIKL (10 days from now) to update PND2 for Client Delay.", client_delay_TIKL_checkbox
                  EditBox 295, 295, 50, 15, verif_req_form_sent_date
                  EditBox 295, 325, 150, 15, worker_signature
                  GroupBox 5, 345, 345, 25, "Dialog Tabs"
                  Text 10, 355, 335, 10, "                       |                    |                   |                   |                    |                    |                     | 8 - Interview"

                  ButtonGroup ButtonPressed
                    PushButton 5, 265, 50, 10, "Verifs needed:", verif_button
                    PushButton 10, 355, 45, 10, "1 - Personal", dlg_one_button
                    PushButton 60, 355, 35, 10, "2 - JOBS", dlg_two_button
                    PushButton 100, 355, 35, 10, "3 - BUSI", dlg_three_button
                    PushButton 140, 355, 35, 10, "4 - CSES", dlg_four_button
                    PushButton 180, 355, 35, 10, "5 - UNEA", dlg_five_button
                    PushButton 220, 355, 35, 10, "6 - Other", dlg_six_button
                    PushButton 260, 355, 40, 10, "7 - Assets", dlg_seven_button
                    PushButton 405, 350, 35, 15, "Done", finish_dlgs_button
                    CancelButton 445, 350, 50, 15
                    OkButton 650, 500, 50, 15
                  Text 5, 15, 55, 10, "Next ER REVW:"
                  Text 280, 15, 50, 10, "* CAF status:"
                  Text 5, 35, 55, 10, "* Actions taken:"
                  ' GroupBox 5, 50, 490, 70, "SNAP Expedited"
                  If the_process_for_snap = "Application" Then
                    If exp_det_case_note_found = False  Then GroupBox 5, 50, 490, 80, "*** SNAP Expedited"
                    If exp_det_case_note_found = True Then GroupBox 5, 50, 490, 80, "SNAP Expedited"

                    If exp_det_case_note_found = TRUE Then
                        Text 15, 60, 400, 10, "EXPEDITED DETERMINATION CASE/NOTE FOUND"
                    Else
                        If full_determination_done = False Then
                            Text 15, 60, 400, 10, "EXPEDITED DETERMINATION NEEDED!!! Press the button below."

                            ButtonGroup ButtonPressed
                              PushButton 340, 110, 150, 15, "Complete Expedited Determination", run_determination_btn
                        Else
                            Text 15, 60, 180, 10, case_assesment_text

                            Text 20, 70, 470, 20, next_steps_one
                            Text 20, 90, 470, 20, next_steps_two
                            Text 20, 110, 320, 20, next_steps_three
                            ' Text 25, 100, 265, 20, next_steps_four
                            ButtonGroup ButtonPressed
                              PushButton 340, 110, 150, 15, "Update Expedited Determination", run_determination_btn
                        End If
                    End If
                  End If

                  '     Text 15, 65, 120, 10, "* Is this SNAP Application Expedited?"
                  '     Text 15, 85, 75, 10, "* EXP Approval Date:"
                  '     Text 195, 65, 75, 10, "* App Month - Income:" '135'
                  '     Text 320, 65, 30, 10, "* Assets:" '260'
                  '     Text 405, 65, 40, 10, "* Expenses:" '345'
                  ' Else
                  ' Text 15, 65, 120, 10, "Is this SNAP Application Expedited?"
                  ' Text 20, 85, 65, 10, "EXP Approval Date:"
                  ' Text 195, 65, 70, 10, "App Month - Income:" '135'
                  ' Text 320, 65, 25, 10, "Assets:" '260'
                  ' Text 405, 65, 40, 10, "Expenses:" '345'
                  ' End If
                  ' Text 135, 50, 90, 10, "CAF Date: " & CAF_datestamp
                  ' Text 135, 85, 55, 10, "Explain Delays:"
                  ' Text 15, 105, 75, 10, "SNAP Denial Date:"
                  ' Text 135, 105, 55, 10, "Explain denial:"
                  GroupBox 5, 130, 490, 105, "Common elements workers should case note:"
                  GroupBox 15, 140, 100, 90, "Application Processing"
                  GroupBox 120, 140, 90, 90, "Form Actions"
                  GroupBox 215, 140, 175, 90, "Interview"
                  Text 5, 245, 50, 10, "Other notes:"
                  GroupBox 5, 280, 280, 60, "Actions the script can do:"
                  Text 295, 285, 120, 10, "Date Verification Request Form Sent:"
                  Text 295, 315, 60, 10, "Worker signature:"
                EndDialog

                Dialog Dialog1
                save_your_work
                cancel_confirmation
                MAXIS_dialog_navigation
                verification_dialog

                If ButtonPressed = tips_and_tricks_xfs_button Then tips_msg = MsgBox("*** Expedited SNAP ***" & vbNewLine & "Anytime the CAF script is run for SNAP at application, expedited information is required. Since the interview is complete, you have enough information to make an EXPEDITED DETERMINATION (different from the screening already completed)." & vbNewLine & vbNewLine &_
                                                                                     "The only time this information is NOT required is if you have run the separate script 'Expedited Determination' - this script will find the case note from that script run and allow you to skip this part of the dialog. The Expedited Determination script has more autofill specific to this process and more detail explained in the note." & vbNewLine & vbNewLine &_
                                                                                     "* App Month Income - Enter the total amount of income received in the month of application here. This income does NOT need to be verified. The script does not caclulate this for you." & vbNewLine & vbNewLine &_
                                                                                     "* App Month Assets - Enter the total LIQUID assets the client has available to them in the month of application. This field is also available on the Asset dialog and will carry over. The script does not calculate this for you." & vbNewLine & vbNewLine &_
                                                                                     "* App Month Expenses - Enter the shelter expense paid/responsible in the month of application plus the standard utility for which the client can claim. If these are completed correctly in the previous dialog, the script will calculate this for you when it first displays. You can change it." & vbNewLine & vbNewLine &_
                                                                                     "THIS INFORMATION SHOULD NOT BE FROM CAF1, but from the client's report and conversation complted during the interview along with any documentation we do have on file - though most are not required." & vbNewLine & vbNewLine &_
                                                                                     "Based on these amounts, enter if the client is expedited or not using the dorpdown with 'Yes' or 'No'. No other consideration should be made to determine the client's eligibility for Expedited. Answering Yes here does not mean you have approved it BUT it does mean the client is eligible for expedited processing." & vbNewLine & vbNewLine &_
                                                                                     "In most situations, the case should be approved if determined to be expedited. If the approval is done or will be done shortly, enter the date of approval." & vbNewLine & vbNewLine &_
                                                                                     "If the approval took more than the expedited processing time (7 days) then explain the delay - this may very well be that no interview had been completed." & vbNewLine & vbNewLine &_
                                                                                     "If the approval cannot be made - leave the date of approval blank and detail what is preventing the approval. Very few things prevent the approval of Expedited SNAP. If you are unsure, check the HSR Manual or contact Knowledge Now.", vbInformation, "Tips and Tricks")

                If ButtonPressed = tips_and_tricks_xfs_button Then ButtonPressed = dlg_eight_button
                If ButtonPressed = -1 Then ButtonPressed = finish_dlgs_button
                If ButtonPressed = verif_button then ButtonPressed = dlg_eight_button
                If ButtonPressed = run_determination_btn Then
                    full_determination_done = True
                    Call run_expedited_determination_script_functionality(xfs_screening, caf_one_income, caf_one_assets, caf_one_rent, caf_one_utilities, determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, approval_date, CAF_datestamp, interview_date, applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, snap_denial_date, snap_denial_explain, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, postponed_verifs_yn, list_postponed_verifs, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, action_due_to_out_of_state_benefits, case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, previous_date_of_application, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn, delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
                End If

                Call assess_button_pressed

                If ButtonPressed = finish_dlgs_button Then
                    'DIALOG 1
                    'New error message formatting for ease of reading.
                    If IsDate(CAF_datestamp) = FALSE Then full_err_msg = full_err_msg & "~!~" & "1^* CAF DATESTAMP ##~##   - Enter a valid date for the CAF datestamp.##~##"

                    For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
                      If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
                          ' MsgBox "Name: " & ALL_MEMBERS_ARRAY(clt_name, the_member) & vbNewLine & "Age: " & ALL_MEMBERS_ARRAY(clt_age, the_member)
                          If ALL_MEMBERS_ARRAY(clt_age, the_member) > 17 OR ALL_MEMBERS_ARRAY(memb_numb, the_member) = "01" Then
                              ALL_MEMBERS_ARRAY(id_detail, the_member) = trim(ALL_MEMBERS_ARRAY(id_detail, the_member))
                              If ALL_MEMBERS_ARRAY(clt_id_verif, the_member) = "OT - Other Document" AND ALL_MEMBERS_ARRAY(id_detail, the_member) = "" Then full_err_msg = full_err_msg & "~!~1^* DETAIL (ID Verif for " & ALL_MEMBERS_ARRAY(clt_name, the_member) & ") ##~##   - Any ID type of OT (Other) needs explanation of what is used for ID verification."
                          End If
                      End If
                    Next
                    If the_process_for_cash = "Application" AND trim(ABPS) <> "" Then
                        If trim(CS_forms_sent_date) <> "N/A" AND IsDate(CS_forms_sent_date) = False AND cash_checkbox = checked Then full_err_msg = full_err_msg & "~!~" & "1^* DATE CS FORMS SENT ##~##   - Enter a valid date for the day that child support forms were sent or given to the client. This is required for Cash cases at application with absent parents.##~##"
                    End If

                    'DIALOG 2
                    For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
                        If ALL_JOBS_PANELS_ARRAY(employer_name, each_job) <> "" THen
                            IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = FALSE Then
                                If ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = unchecked Then
                                    ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) = trim(ALL_JOBS_PANELS_ARRAY(budget_explain, each_job))
                                    If ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) = "" Then
                                        full_err_msg = full_err_msg & "~!~" & "2^* EXPLAIN BUDGET for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - Additional detail about how the job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " - was budgeted is required. Complete the 'Explain Budget' field for this job."
                                    ElseIf len(ALL_JOBS_PANELS_ARRAY(budget_explain, each_job)) < 20 Then
                                        full_err_msg = full_err_msg & "~!~" & "2^* EXPLAIN BUDGET for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - Budget detail for job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " - should be longer. Budget cannot be sufficiently explained in a short note."
                                    End If
                                End If
                                If SNAP_checkbox = checked Then
                                    If IsNumeric(ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job)) = FALSE Then full_err_msg = full_err_msg & "~!~" & "2^* SNAP PIC - PAY DATE AMOUNT for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - For a SNAP case the average pay date amount must be entered as a number. Update the 'Pay Date Amount' for job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & "."
                                    If ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job) = "Type or select" Then full_err_msg = full_err_msg & "~!~" & "2^* SNAP PIC - PAY FREQUENCY for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - The pay frequency for SNAP pay date amount needs to be identified to correctly note the income. Update the frequency after 'Pay Date Amount' for the job - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & "."
                                    If IsNumeric(ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job)) = False Then full_err_msg = full_err_msg & "~!~" & "2^* SNAP PIC - PROSPECTIVE AMOUNT for " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & " ##~##   - For SNAP cases, the monthly prospective amount needs to be entered as a number in the 'Prospective Amount' field for jobw - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job) & "."
                                End If
                            End If
                        End If
                    Next

                    'DIALOG 3
                    If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then
                        For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
                            ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) = trim(ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi))
                            If ALL_BUSI_PANELS_ARRAY(estimate_only, each_busi) = unchecked Then
                                If ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) = "" Then
                                    full_err_msg = full_err_msg & "~!~3^* EXPLAIN BUDGET for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Additional detail about how BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " was budgeted is required. Complete the 'Explain Budget' field for this self employment."
                                ElseIf len(ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi)) < 20 Then
                                    full_err_msg = full_err_msg & "~!~3^* EXPLAIN BUDGET for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Additional detail about how BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " was budgeted should be longer - the note is too short so sufficiently explain how the income was budgeted."
                                End If
                            End If
                            If ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) = "Select One" Then full_err_msg = full_err_msg & "~!~3^* SELF EMPLOYMENT METHOD for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Indicate which calculation method will be used for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & "."
                            If ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) = "Tax Forms" Then
                                If SNAP_checkbox = checked Then
                                    If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "Income Tax Returns" AND trim(ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi)) = "" Then full_err_msg = full_err_msg & "~!~3^* EXPENSES NOT ALLOWED for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Since the calculation method is 'Tax Forms' and this is a SNAP case with Tax Forms verifying, indicate what (if any) expenses on taxes have been excluded."
                                    If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "Pend Out State Verif" OR ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "No Verif Provided" OR ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) = "Delayed Verif" Then
                                    Else
                                        If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* SNAP INCOME VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of income for SNAP should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                        If ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* SNAP EXPENSE VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of expenses for SNAP should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                    End If
                                End If
                                If cash_checkbox = checked or EMER_checkbox = checked Then
                                    If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) = "Pend Out State Verif" OR ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) = "No Verif Provided" OR ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) = "Delayed Verif" Then
                                    Else
                                        If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* CASH INCOMME VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of income for Cash/EMER should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                        If ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "Income Tax Returns" Then full_err_msg = full_err_msg & "~!~3^* CASH EXPENSE VERIFICATION for BUSI " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " " & ALL_BUSI_PANELS_ARRAY(panel_instance, each_busi) & " ##~##   - Verification of expenses for Cash/EMER should be 'Income Tax Returns' when calculation method is 'Tax Forms'"
                                    End If
                                End If
                            End If
                        Next
                    End If

                    'DIALOG 4

                    'DIALOG 5
                    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                        If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
                            If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)) <> "" AND trim(UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)) = "" Then full_err_msg = full_err_msg & "~!~5^* RSDI NOTES for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Explain details about RSDI Income and Budgeting."
                            If trim(UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)) <> "" AND trim(UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)) = "" Then full_err_msg = full_err_msg & "~!~5^* SSI NOTES for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Explain details about SSI Income and Budgeting."
                        End If
                    Next
                    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
                        If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
                            If SNAP_checkbox = checked and IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)) = False Then full_err_msg = full_err_msg & "~!~5^* UC SNAP PROSP AMOUNT for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Indicate the prospective amount of UC income that will be budgeted for SNAP."
                            If UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb) <> "" Then
                                If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = False Then
                                    full_err_msg = full_err_msg & "~!~5^* UC TIKL DATE for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - In order to set a TIKL, a valid date needs to be entered in the box for the UC TIKL."
                                Else
                                    If DateDiff("d", date, UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) < 0 Then full_err_msg = full_err_msg & "~!~5^* UC TIKL DATE for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - To set a TIKL for the end of UC income, the TIKL date must be in the future."
                                End If
                            End If
                            If trim(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) <> "" Then
                                If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) = False Then full_err_msg = full_err_msg & "~!~5^* UC WEEKLY GROSS for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - The UC Gross weekly amount needs to be entered as a number."
                                If UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) <> UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) Then
                                    If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) = FALSE or IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)) = FALSE or IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) = FALSE Then
                                        If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)) = FALSE Then full_err_msg = full_err_msg & "~!~5^* UC BUDGETED WEEKLY AMOUNT for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Enter the UC weekly Net Amount as a number."
                                        If IsNumeric(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) = FALSE Then full_err_msg = full_err_msg & "~!~5^* UC WEEKLY ALLOWED DEDUCTIONS for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Enter the weekly allowed deductions for UC as a number."
                                    Else
                                        calculated_net_weekly = UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) - UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)
                                        UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) = UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) * 1
                                        'MsgBox "Calc Net Weekly - " & calculated_net_weekly & vbCR & "Entered Net Weekly - " & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb)
                                        If calculated_net_weekly <> UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) Then full_err_msg = full_err_msg & "~!~5^* UC WEEKLY GROSS, BUDGETED AMOUNT, ALLOWED DEDUCTIONS for MEMB " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " ##~##   - Review your UC weekly gross, net and counted deductions. The net amount is not equal to the gross amount less counted deductions. ##~## Weekly Gross ($" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ") - Allowed Deductions ($" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ") = $" & calculated_net_weekly & " ##~## Weekly Budgeted Amount $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) &  " ##~## Difference between gross and allowed deductions should equal the weekly budgeted ammount."
                                    End If
                                End If
                            End If
                        End If
                    Next

                    'DIALOG 6
                    If SNAP_checkbox = checked and trim(notes_on_wreg) = "" Then full_err_msg = full_err_msg & "~!~6^* WREG Notes ##~##   - Update WREG detail as this is a SNAP case."
                    ' Removing reqqquirement of living situation because this should be handled during interview.
                    ' If living_situation = "Blank" or living_situation = "  " Then full_err_msg = full_err_msg & "~!~6^* LIVING SITUATION ##~##   - Living situation needs to be entered for each case. 'Blank' is not valid."
                    'We are not erroring for if ADDR verification is 'NO' or '?' - if we get additional policy information that this is necessary - add it here

                    'DIALOG 7
                    ' If SNAP_checkbox = checked and CAF_type = "Application" Then
                    '     If trim(app_month_assets) = "" OR IsNumeric(app_month_assets) = FALSE AND exp_det_case_note_found = FALSE Then full_err_msg = full_err_msg & "~!~7^* Indicate the total of liquid assets in the application month."
                    ' End If

                    If family_cash = TRUE and trim(notes_on_time) = "" Then full_err_msg = full_err_msg & "~!~7^* TIME ##~##   - For a family cash case, detail on TIME needs to be added."
                    If family_cash = TRUE and trim(notes_on_sanction) = "" Then full_err_msg = full_err_msg & "~!~7^* SANCTION ##~##   - This is a family cash case, sanction detail needs to be added."
                    If family_cash = TRUE and trim(EMPS) = "" Then full_err_msg = full_err_msg & "~!~7^* EMMPS ##~##   - EMPS detail needs to be added for a family cash case. "
                    If cash_checkbox = unchecked AND trim(DIET) <> "" Then full_err_msg = full_err_msg & "~!~7^* DIET ##~##   - DIET information should not be entered into a non-cash case."
                    If InStr(shelter_details, "Mortgage") AND trim(notes_on_rest) = "" Then full_err_msg = full_err_msg & "~!~7^* REST ##~##   - SHEL indicates that Mortgage is being paid, but no information has been added to REST. Update Shelter information or add detail to REST."

                    'DIALOG 8
                    If CAF_status = "Select or Type" Then full_err_msg = full_err_msg & "~!~8^* CAF STATUS ##~##   - Indicate the CAF Status."
                    If the_process_for_snap = "Application" AND exp_det_case_note_found = FALSE Then
                        If full_determination_done = False Then full_err_msg = full_err_msg & "~!~8^* COMPLETE EXPEDITED DETERMINATION ##~##   - This is a a SNAP case at application. We must complete the expedited determination. Press the button labeled 'COMPLETE EXPEDITED DETERMINATION' and complete all steps of this functionality to create an expedited determination."
                    End If
                    If trim(actions_taken) = "" Then full_err_msg = full_err_msg & "~!~8^* ACTIONS TAKEN ##~##   - Indicate what actions were taken when processing this CAF."
                    prev_err_msg = full_err_msg
                End If

                Call display_errors(full_err_msg, TRUE)
                If full_err_msg = "" and ButtonPressed = finish_dlgs_button Then pass_eight = true
                If ButtonPressed = finish_dlgs_button Then ButtonPressed = -1
            End If
            ' MsgBox "Button - " & ButtonPressed & vbNewLine & "Pass Eight - " & pass_eight
        Loop until pass_eight = true
        ' MsgBox "Now we call proceed confirmation"
        CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
    Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
    ' MsgBox "Now We call check for password"
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

Call back_to_SELF
If continue_in_inquiry = "" Then
    Do
        Call back_to_SELF
        EMReadScreen MX_region, 12, 22, 48
        MX_region = trim(MX_region)
        If MX_region = "INQUIRY DB" Then

            Dialog1 = ""
            BeginDialog dialog1, 0, 0, 266, 120, "Still in Inquiry"
              ButtonGroup ButtonPressed
                PushButton 165, 80, 95, 15, "Stop the Script Run (ESC)", stop_script_button
                PushButton 140, 100, 120, 15, "Continue - I have switched (Enter)", continue_script
              Text 10, 10, 110, 20, "It appears you are now running in INQUIRY on this session."
              Text 10, 40, 105, 20, "The script cannot update or CASE/NOTE in INQUIRY."
              Text 10, 65, 255, 10, "Switch to Production now to ensure the note is entered and continue the script."
            EndDialog

            Do
                dialog dialog1
                If ButtonPressed = stop_script_button Then ButtonPressed = 0
                If ButtonPressed = 0 Then script_end_procedure("Script ended since it was started in Inquiry.")
                If ButtonPressed = -1 Then ButtonPressed = continue_script

                Call check_for_password(are_we_passworded_out)
            Loop until are_we_passworded_out = FALSE

        Else
            ButtonPressed = continue_script
        End If
    Loop until ButtonPressed = continue_script AND MX_region <> "INQUIRY DB"
End If

If trim(CS_forms_sent_date) = "N/A" Then CS_forms_sent_date = ""

'Go to ADDR to update living situation
Call navigate_to_MAXIS_screen("STAT", "ADDR")
EMReadScreen panel_living_sit, 2, 11, 43
If living_situation = "Blank" or living_situation = "  " Then
    dialog_liv_sit_code = "__"
Else
    dialog_liv_sit_code = left(living_situation, 2)
End If

If dialog_liv_sit_code <> panel_living_sit OR dialog_liv_sit_code = "__" Then
    PF9
    EMWriteScreen dialog_liv_sit_code, 11, 43
    transmit
    EmReadscreen addr_error, 21, 24, 2
    If addr_error = "ONLY ONE FUTURE PANEL" then transmit   'error message that needs to be bypassed if other changes occur in that footer month/year.
End If

Do
    Do
        qual_err_msg = ""

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 451, 205, "CAF Qualifying Questions"
          DropListBox 220, 40, 35, 45, "No"+chr(9)+"Yes", qual_question_one
          ComboBox 330, 40, 115, 45, verification_memb_list, qual_memb_one
          DropListBox 220, 80, 35, 45, "No"+chr(9)+"Yes", qual_question_two
          ComboBox 330, 80, 115, 45, verification_memb_list, qual_memb_two
          DropListBox 220, 110, 35, 45, "No"+chr(9)+"Yes", qual_question_three
          ComboBox 330, 110, 115, 45, verification_memb_list, qual_memb_three
          DropListBox 220, 140, 35, 45, "No"+chr(9)+"Yes", qual_question_four
          ComboBox 330, 140, 115, 45, verification_memb_list, qual_memb_four
          DropListBox 220, 160, 35, 45, "No"+chr(9)+"Yes", qual_question_five
          ComboBox 330, 160, 115, 45, verification_memb_list, qual_memb_five
          ButtonGroup ButtonPressed
            OkButton 340, 185, 50, 15
            CancelButton 395, 185, 50, 15
          Text 10, 10, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the client. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
          Text 10, 40, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
          Text 10, 80, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
          Text 10, 110, 195, 30, "Is anyone in your householdhiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
          Text 10, 140, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
          Text 10, 160, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
          Text 260, 40, 70, 10, "Household Member:"
          Text 260, 80, 70, 10, "Household Member:"
          Text 260, 110, 70, 10, "Household Member:"
          Text 260, 140, 70, 10, "Household Member:"
          Text 260, 160, 70, 10, "Household Member:"
        EndDialog

        dialog Dialog1
        cancel_confirmation

        If qual_question_one = "Yes" AND (trim(qual_memb_one) = "" OR qual_memb_one = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 1, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_two = "Yes" AND (trim(qual_memb_two) = "" OR qual_memb_two = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 2, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_three = "Yes" AND (trim(qual_memb_three) = "" OR qual_memb_three = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 3, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_four = "Yes" AND (trim(qual_memb_four) = "" OR qual_memb_four = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 4, yes is indicated however no member is listed - please enter the member that this question applies to."
        If qual_question_five = "Yes" AND (trim(qual_memb_five) = "" OR qual_memb_five = "Select or Type") Then qual_err_msg = qual_err_msg & vbNewLine & "* For Quesion 5, yes is indicated however no member is listed - please enter the member that this question applies to."

        If qual_err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & qual_err_msg
    Loop until qual_err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

qual_questions_yes = FALSE
If qual_question_one = "Yes" Then qual_questions_yes = TRUE
If qual_question_two = "Yes" Then qual_questions_yes = TRUE
If qual_question_three = "Yes" Then qual_questions_yes = TRUE
If qual_question_four = "Yes" Then qual_questions_yes = TRUE
If qual_question_five = "Yes" Then qual_questions_yes = TRUE

'Now, the client_delay_checkbox business. It'll update client delay if the box is checked and it isn't a recert.
If client_delay_checkbox = checked and application_processing = True then
	call navigate_to_MAXIS_screen("REPT", "PND2")

    limit_reached = FALSE
    row = 1
    col = 1
    EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
    If row <> 0 Then
        transmit
        limit_reached = TRUE
    End If

    If limit_reached = TRUE Then
        PND2_row = 7
        Do
            EMReadScreen PND2_case_number, 8, PND2_row, 5
            if trim(PND2_case_number) = MAXIS_case_number Then Exit Do
            PND2_row = PND2_row + 1
        Loop until PND2_row = 18
    Else
        EMGetCursor PND2_row, PND2_col
    End If

    If PND2_row = 18 Then
        client_delay_checkbox = unchecked
        MsgBox "The script could not navigate to REPT/PND2 due to a MAXIS display limit. This case will not be updated for client delay. Please email to BlueZone Script Team with the case number and report that the Display Limit on REPT/PND2 was reached."
    End If

	for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
		EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
		If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
		EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
		If PND2_HC_status_check = "P" then
			EMWriteScreen "X", PND2_row, 3
			transmit
			person_delay_row = 7
			Do
				EMReadScreen person_delay_check, 1, person_delay_row, 39
				If person_delay_check <> " " then EMWriteScreen "C", person_delay_row, 39
				person_delay_row = person_delay_row + 2
			Loop until person_delay_check = " " or person_delay_row > 20
			PF3
		End if
		EMReadScreen additional_app_check, 14, PND2_row + 1, 17
		If additional_app_check <> "ADDITIONAL APP" then exit for
		PND2_row = PND2_row + 1
	next
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_checkbox = unchecked		'Probably unnecessary except that it changes the case note parameters
	End if
End if

'Going to TIKL. Now using the write TIKL function
If TIKL_checkbox = checked and application_processing = True then
	If DateDiff ("d", CAF_datestamp, date) > 30 Then 'Error handling to prevent script from attempting to write a TIKL in the past
		MsgBox "Cannot set TIKL as CAF Date is over 30 days old and TIKL would be in the past. You must manually track."
        TIKL_checkbox = unchecked
	Else
        If cash_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "Cash/"
        If GRH_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "GRH/"
        If SNAP_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "SNAP/"
        If EMER_checkbox = checked then TIKL_msg_one = TIKL_msg_one & "EMER/"
        TIKL_msg_one = Left(TIKL_msg_one, (len(TIKL_msg_one) - 1))
        TIKL_msg_one = TIKL_msg_one & " has been pending for 30 days. Evaluate for possible denial."
		'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
        Call create_TIKL(TIKL_msg_one, 30, CAF_datestamp, False, TIKL_note_text)
        Call back_to_SELF
	End If
ElseIf TIKL_checkbox = checked and application_processing = False then
    TIKL_checkbox = unchecked
End if
If client_delay_TIKL_checkbox = checked then
    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL(">>>UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE<<<", 10, date, False, TIKL_note_text)
    Call back_to_SELF
End if

For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
        If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = TRUE Then
            'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
            tikl_msg = "Review UC Income for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " as it may have ended or be near ending."
            Call create_TIKL(TIKL_msg, 10, UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb), False, TIKL_note_text)
            Call back_to_SELF
        End If
    End If
Next
'--------------------END OF TIKL BUSINESS

If HC_checkbox = checked Then
    call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
    If the_process_for_hc = "Application" Then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", HC_datestamp)
    If the_process_for_hc = "Recertification" Then call autofill_editbox_from_MAXIS(HH_member_array, "REVW", HC_datestamp)
    call autofill_editbox_from_MAXIS(HH_member_array, "ACCI", hc_acci_info)
    call autofill_editbox_from_MAXIS(HH_member_array, "BILS", hc_bils_info)
    call autofill_editbox_from_MAXIS(HH_member_array, "FACI", hc_faci_info)
    call autofill_editbox_from_MAXIS(HH_member_array, "INSA", hc_insa_info)
    If CAF_form = "Combined AR for Certain Pops (DHS-3727)" Then
        HC_document_received = "DHS-3727 (Combined AR for Certain Pops)"
        HC_datestamp = CAF_datestamp & ""
    End If
    hc_medi_info = MEDI
    hc_faci_info = FACI

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 481, 295, "HC Detail"
      ComboBox 80, 5, 150, 15, "Select or Type"+chr(9)+"DHS-2128 (LTC Renewal)"+chr(9)+"DHS-3417B (Req. to Apply...)"+chr(9)+"DHS-3418 (HC Renewal)"+chr(9)+"DHS-3531 (LTC Application)"+chr(9)+"DHS-3876 (Certain Pops App)"+chr(9)+"DHS-6696 (MNsure HC App)"+chr(9)+"DHS-3727 (Combined AR for Certain Pops)"+chr(9)+HC_document_received, HC_document_received
      EditBox 80, 20, 50, 15, HC_datestamp
      ComboBox 360, 5, 115, 15, "Select of Type"+chr(9)+"incomplete"+chr(9)+"approved"+chr(9)+"denied"+chr(9)+HC_form_status, HC_form_status
      CheckBox 310, 25, 80, 10, "Application signed?", HC_application_signed_check
      CheckBox 405, 25, 65, 10, "MMIS updated?", MMIS_updated_check
      EditBox 65, 40, 165, 15, retro_request
      EditBox 290, 40, 185, 15, hc_hh_comp
      EditBox 35, 60, 440, 15, hc_medi_info
      EditBox 35, 80, 440, 15, hc_insa_info
      EditBox 35, 100, 440, 15, hc_acci_info
      EditBox 35, 120, 440, 15, hc_bils_info
      EditBox 35, 140, 440, 15, hc_faci_info
      EditBox 55, 160, 420, 15, waiver_ltc_info
      EditBox 55, 180, 420, 15, spenddown_info
      CheckBox 55, 200, 245, 10, "Check here to have the script create a TIKL to deny at the 45 day mark.", hc_tikl_checkbox
      EditBox 55, 215, 420, 15, hc_other_notes
      EditBox 55, 235, 420, 15, hc_verifs_needed
      EditBox 55, 255, 420, 15, hc_actions_taken
      ButtonGroup ButtonPressed
        OkButton 370, 275, 50, 15
        CancelButton 425, 275, 50, 15
        PushButton 5, 65, 25, 10, "MEDI:", MEDI_button
        PushButton 5, 85, 25, 10, "INSA", INSA_button
        PushButton 5, 105, 25, 10, "ACCI:", ACCI_button
        PushButton 5, 125, 25, 10, "BILS:", BILS_button
        PushButton 5, 145, 25, 10, "FACI:", FACI_button
      Text 10, 10, 70, 10, "HC Form Received:"
      Text 30, 25, 45, 10, "Date Stamp:"
      Text 300, 10, 60, 10, "HC Form status:"
      Text 10, 45, 50, 10, "Retro Request:"
      Text 240, 45, 50, 10, "HC HH Comp:"
      Text 10, 165, 45, 10, "Waiver/LTC:"
      Text 10, 185, 40, 10, "Spenddown:"
      Text 10, 220, 40, 10, "Other notes:"
      Text 5, 240, 50, 10, "Verifs needed:"
      Text 5, 260, 50, 10, "Actions taken:"
    EndDialog

    Do
        Do
            hc_err_msg = ""

            dialog dialog1

            cancel_confirmation
            MAXIS_dialog_navigation

            If IsDate(HC_datestamp) = False Then hc_err_msg = hc_err_msg & vbNewLine & "* Enter the date of the form as a valid date."
            If trim(HC_form_status) = "" OR trim(HC_form_status) = "Select or Type" Then hc_err_msg = hc_err_msg & vbNewLine & "* Indicate (by selection or typing manually) the status of the health care form being processed."
            If trim(HC_document_received) = "" or trim(HC_document_received) = "Select or Type" Then hc_err_msg = hc_err_msg & vbNewLine & "* Indicate (by selection or typing manually) what form is being processed for health care."

            If hc_err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & hc_err_msg

        Loop until hc_err_msg = ""
        call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If hc_tikl_checkbox = checked Then
        If DateDiff ("d", HC_datestamp, date) > 45 Then 'Error handling to prevent script from attempting to write a TIKL in the past
            MsgBox "Cannot set TIKL as HC Form Date is over 45 days old and TIKL would be in the past. You must manually track."
            hc_tikl_checkbox = unchecked
        Else
            'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
            Call create_TIKL("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.", 45, HC_datestamp, False, TIKL_note_text)
            Call back_to_SELF
        End If
    End If
End If

'Adding footer month to the recertification case notes
' If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " recert"
progs_list = ""
If cash_checkbox = checked Then progs_list = progs_list & ", Cash"
If GRH_checkbox = checked Then progs_list = progs_list & ", GRH"
If SNAP_checkbox = checked Then progs_list = progs_list & ", SNAP"
If EMER_checkbox = checked Then progs_list = progs_list & ", EMER"
If left(progs_list, 1) = "," Then progs_list = right(progs_list, len(progs_list) - 2)

prog_and_type_list = ""
If cash_checkbox = checked Then
    If the_process_for_cash = "Application" Then prog_and_type_list = prog_and_type_list & ", Cash App"
    If the_process_for_cash = "Recertification" Then prog_and_type_list = prog_and_type_list & ", " & cash_recert_mo & "/" & cash_recert_yr & " Cash Recert"
End If
If GRH_checkbox = checked Then
    If the_process_for_grh = "Application" Then prog_and_type_list = prog_and_type_list & ", GRH App"
    If the_process_for_grh = "Recertification" Then prog_and_type_list = prog_and_type_list & ", " & grh_recert_mo & "/" & grh_recert_yr & " GRH Recert"
End If
If snap_checkbox = checked Then
    If the_process_for_snap = "Application" Then prog_and_type_list = prog_and_type_list & ", SNAP App"
    If the_process_for_snap = "Recertification" Then prog_and_type_list = prog_and_type_list & ", " & snap_recert_mo & "/" & snap_recert_yr & " SNAP Recert"
End If
If EMER_checkbox = checked Then prog_and_type_list = prog_and_type_list & ", EMER App"
If left(prog_and_type_list, 1) = "," Then prog_and_type_list = right(prog_and_type_list, len(prog_and_type_list) - 2)

If SNAP_checkbox = checked Then
    adult_snap_count = adult_snap_count * 1
    child_snap_count = child_snap_count * 1
    total_snap_count = adult_snap_count + child_snap_count
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_snap_checkbox, each_member) = checked Then
            included_snap_members = included_snap_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        Else
            If ALL_MEMBERS_ARRAY(count_snap_checkbox, each_member) = checked Then counted_snap_members = counted_snap_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        End If
    Next
    If included_snap_members <> "" Then included_snap_members = right(included_snap_members, len(included_snap_members) - 2)
    If counted_snap_members <> "" Then counted_snap_members = right(counted_snap_members, len(counted_snap_members) - 2)
End If
If cash_checkbox = checked Then
    adult_cash_count = adult_cash_count * 1
    child_cash_count = child_cash_count * 1
    total_cash_count = adult_cash_count + child_cash_count
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_cash_checkbox, each_member) = checked Then
            included_cash_members = included_cash_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        Else
            If ALL_MEMBERS_ARRAY(count_cash_checkbox, each_member) = checked Then counted_cash_members = counted_cash_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        End If
    Next
    If included_cash_members <> "" Then included_cash_members = right(included_cash_members, len(included_cash_members) - 2)
    If counted_cash_members <> "" Then counted_cash_members = right(counted_cash_members, len(counted_cash_members) - 2)
End If
If EMER_checkbox = checked Then
    adult_emer_count = adult_emer_count * 1
    child_emer_count = child_emer_count * 1
    total_emer_count = adult_emer_count + child_emer_count
    For each_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
        If ALL_MEMBERS_ARRAY(include_emer_checkbox, each_member) = checked Then
            included_emer_members = included_emer_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        Else
            If ALL_MEMBERS_ARRAY(count_emer_checkbox, each_member) = checked Then counted_emer_members = counted_emer_members & ", M" & ALL_MEMBERS_ARRAY(memb_numb, each_member)
        End If
    Next
    If included_emer_members <> "" Then included_emer_members = right(included_emer_members, len(included_emer_members) - 2)
    If counted_emer_members <> "" Then counted_emer_members = right(counted_emer_members, len(counted_emer_members) - 2)
End If

'Determining if there are
'Income
case_has_income = FALSE

If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then case_has_income = TRUE
If trim(notes_on_jobs) <> "" Then case_has_income = TRUE
If trim(earned_income) <> "" Then case_has_income = TRUE
If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then case_has_income = TRUE
If trim(notes_on_busi) <> "" Then case_has_income = TRUE
For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then case_has_income = TRUE
    If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then case_has_income = TRUE
    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then case_has_income = TRUE
Next
If trim(notes_on_cses) <> "" Then case_has_income = TRUE
If trim(notes_on_ssa_income) <> "" Then case_has_income = TRUE
If trim(notes_on_VA_income) <> "" Then case_has_income = TRUE
If trim(notes_on_WC_income) <> "" Then case_has_income = TRUE
If trim(other_uc_income_notes) <> "" Then case_has_income = TRUE
If trim(notes_on_other_UNEA) <> "" Then case_has_income = TRUE

'Personal
case_has_personal = FALSE

If trim(cit_id) <> "" Then case_has_personal = TRUE
If trim(IMIG) <> "" Then case_has_personal = TRUE
If trim(SCHL) <> "" Then case_has_personal = TRUE
If trim(DISA) <> "" Then case_has_personal = TRUE
If trim(FACI) <> "" Then case_has_personal = TRUE
If trim(PREG) <> "" Then case_has_personal = TRUE
If trim(ABPS) <> "" Then case_has_personal = TRUE
If CS_forms_sent_date <> "" Then case_has_personal = TRUE
If trim(AREP) <> "" Then case_has_personal = TRUE
If address_confirmation_checkbox = checked Then case_has_personal = TRUE
If homeless_yn = "Yes" Then case_has_personal = TRUE
If trim(addr_county) <> "" Then case_has_personal = TRUE
If trim(living_situation) <> "" Then case_has_personal = TRUE
If trim(notes_on_address) <> "" Then case_has_personal = TRUE
If trim(DISQ) <> "" Then case_has_personal = TRUE
If trim(notes_on_wreg) <> "" Then case_has_personal = TRUE
all_abawd_notes = notes_on_abawd & notes_on_abawd_two & notes_on_abawd_three
If trim(all_abawd_notes) <> "" Then case_has_personal = TRUE
If trim(notes_on_time) <> "" Then case_has_personal = TRUE
If trim(notes_on_sanction) <> "" Then case_has_personal = TRUE
If trim(EMPS) <> "" Then case_has_personal = TRUE
If trim(MEDI) <> "" Then case_has_personal = TRUE
If trim(DIET) <> "" Then case_has_personal = TRUE
If trim(case_changes) <> "" Then case_has_personal = TRUE

'Resources
case_has_resources = FALSE

If confirm_no_account_panel_checkbox = checked Then case_has_resources = TRUE
If trim(notes_on_acct) <> "" Then case_has_resources = TRUE
If trim(notes_on_cash) <> "" Then case_has_resources = TRUE
If trim(notes_on_cars) <> "" Then case_has_resources = TRUE
If trim(notes_on_rest) <> "" Then case_has_resources = TRUE
If trim(notes_on_other_assets) <> "" Then case_has_resources = TRUE

'Expenses
case_has_expenses = FALSE

If trim(total_shelter_amount) <> "" Then case_has_expenses = TRUE
If trim(full_shelter_details) <> "" Then case_has_expenses = TRUE
If trim(notes_on_acut) <> "" Then case_has_expenses = TRUE
If hest_information <> "Select ALLOWED HEST" Then case_has_expenses = TRUE
If trim(notes_on_coex) <> "" Then case_has_expenses = TRUE
If trim(notes_on_dcex) <> "" Then case_has_expenses = TRUE
If trim(notes_on_other_deduction) <> "" Then case_has_expenses = TRUE
If trim(expense_notes) <> "" Then case_has_expenses = TRUE
If trim(FMED) <> "" Then case_has_expenses = TRUE

'THE CASE NOTES-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Expedited Determination Case Note
'Navigates to case note, and checks to make sure we aren't in inquiry.
case_notes_information = "CASE NOTES ATTEMPTED AND OUTCOME %^% %^%"
If HC_checkbox = unchecked Then case_notes_information = case_notes_information & "No HC NOTE Attempted - HC not checked %^% %^%"
If HC_checkbox = checked Then
    case_notes_information = case_notes_information & "HC NOTE Attempted %^%"
    hc_note_header = HC_datestamp & " " & HC_document_received & ": " & HC_form_status
    case_notes_information = case_notes_information & "Script Header - " & hc_note_header & " %^%"

    Call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE(hc_note_header)

    Call write_bullet_and_variable_in_CASE_NOTE("Actions Taken", hc_actions_taken)

    If HC_application_signed_check = checked Then Call write_variable_in_CASE_NOTE("* HC form was signed.")

    Call write_bullet_and_variable_in_CASE_NOTE("HC form received", HC_datestamp)
    Call write_bullet_and_variable_in_CASE_NOTE("Retro Request", retro_request)
    Call write_bullet_and_variable_in_CASE_NOTE("HC HH Comp", hc_hh_comp)
    Call write_bullet_and_variable_in_CASE_NOTE("Spenddown", spenddown_info)

    'INCOME
    If case_has_income = TRUE Then
        Call write_variable_in_CASE_NOTE("===== INCOME =====")
    Else
        Call write_variable_in_CASE_NOTE("== No Income detail Listed for this case. ==")
    End If
    'JOBS
    If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then
        ' Call write_variable_with_indent(variable_name)
        Call write_variable_in_CASE_NOTE("--- JOBS Income ---")
        For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
            Call write_variable_in_CASE_NOTE("Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " at " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job))
            If ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = checked Then Call write_variable_in_CASE_NOTE("* This job has not been verified and this is only an estimate.")
            IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then call write_variable_in_CASE_NOTE("* BUDGET DETAIL ABOUT THIS JOB IN PREVIOUS CASE NOTE.")
            If ALL_JOBS_PANELS_ARRAY(verif_code, each_job) = "Delayed" Then
                Call write_variable_in_CASE_NOTE("* Verification of this job has been delayed for review or approval of Expedited SNAP.")
            ElseIf ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = unchecked Then
                Call write_variable_in_CASE_NOTE("* Verification - " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job))
            End If
            Call write_bullet_and_variable_in_CASE_NOTE("Verification", ALL_JOBS_PANELS_ARRAY(verif_explain, each_job))
            If ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Retro Income: $" & ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(retro_hours, each_job) & " hours.")
            If ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Prospective Income: $" & ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job) & " hours.")
            If ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("About Budget: " & ALL_JOBS_PANELS_ARRAY(budget_explain, each_job))
        Next
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("JOBS", notes_on_jobs)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Earned Income", earned_income)

    'BUSI
    If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then
        Call write_variable_in_CASE_NOTE("--- BUSI Income ---")
        For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
            busi_det_msg = "Self Employment for Memb " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " - BUSI type:" & right(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi), len(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi)) - 4) & "."
            Call write_variable_in_CASE_NOTE(busi_det_msg)

            If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Description: " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) & ".")
            If ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business structure: " & ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) & ".")
            If ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Clt owns " & ALL_BUSI_PANELS_ARRAY(share_num, each_busi) & "/" & ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) & " of the business.")
            If ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business also owned by Memb(s) " & ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) & ".")

            se_method_det_msg = "* Self Employment Budgeting method selected: " & ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) & "."
            Call write_variable_in_CASE_NOTE(se_method_det_msg)
            If ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Method selected on: " & ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) & ".")
            If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, each_busi) = checked Then Call write_variable_with_indent_in_CASE_NOTE("The self employment method selected was discussed with the client.")

            If cash_checkbox = checked OR EMER_checkbox = checked Then
                Call write_variable_in_CASE_NOTE("* Cash Income and Expense Detail:")
                cash_income_det = ""
                cash_expense_det = ""

                If ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) <> "" Then
                    cash_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) & " - "
                    cash_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_cash, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) <> "" Then
                    cash_income_det = cash_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) & " - "
                    cash_expense_det = cash_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_cash, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "" Then cash_income_det = cash_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi)
                If ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "" Then cash_expense_det = cash_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi)

                Call write_variable_with_indent_in_CASE_NOTE(cash_income_det)
                Call write_variable_with_indent_in_CASE_NOTE(cash_expense_det)
            End If
            If SNAP_checkbox = checked Then
                Call write_variable_in_CASE_NOTE("* SNAP Income and Expense Detail:")
                snap_income_det = ""
                snap_expense_det = ""

                If ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) <> "" Then
                    snap_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) & " - "
                    snap_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_snap, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) <> "" Then
                    snap_income_det = snap_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) & " - "
                    snap_expense_det = snap_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi) & " - "
                End If
                If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "" Then snap_income_det = snap_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi)
                If ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "" Then snap_expense_det = snap_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi)

                Call write_variable_with_indent_in_CASE_NOTE(snap_income_det)
                Call write_variable_with_indent_in_CASE_NOTE(snap_expense_det)
                If ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Expenses from taxes not allowed: " & ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi))
            End If
            rept_hours_det_msg = ""
            min_wg_hours_det_msg = ""
            If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = ""
            If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = ""
            If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = ""
            If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = ""

            If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then
                rept_hours_det_msg = rept_hours_det_msg & "Clt reported monthly work hours of: "
                If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) & " retrospecive work and "
                If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) & " prospoective work hrs"
                rept_hours_det_msg = rept_hours_det_msg & ". "
            End If
            If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then
                min_wg_hours_det_msg = min_wg_hours_det_msg & "Work earnings indicate Minumun Wage Hours of: "
                If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) & " retrospective and "
                If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) & " prospective"
                min_wg_hours_det_msg = min_wg_hours_det_msg & ". "
            End If
            If rept_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & rept_hours_det_msg)
            If min_wg_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & min_wg_hours_det_msg)
            If ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Verif Detail: " & ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi))
            If ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Budget Detail: " & ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi))
        Next
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("BUSI", notes_on_busi)

    'CSES
    If show_cses_detail = TRUE Then
        Call write_variable_in_CASE_NOTE("--- Child Support Income ---")
        For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
            If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
                total_cs = 0
                If IsNumeric(UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)

                Call write_variable_in_CASE_NOTE("* Total child support income for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & ": $" & total_cs)
                If UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb) <> "" Then
                    cs_disb_inc_det = "Disbursed child support: $" & UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)) <> "" Then cs_disb_inc_det = cs_disb_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)

                    Call write_variable_with_indent_in_CASE_NOTE(cs_disb_inc_det)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb) & " month(s) of disbursement income.")
                    If trim(UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb))
                End If

                If UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb) <> "" Then
                    cs_arrears_inc_det = "Disbursed child support arrears: $" & UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)) <> "" Then cs_arrears_inc_det = cs_arrears_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)

                    Call write_variable_with_indent_in_CASE_NOTE(cs_arrears_inc_det)
                    If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb) & " month(s) of disbursement income.")
                    If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb))
                End If

                If UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb) <> "" Then
                    direct_cs_inc_det = "Direct child support: $" & UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                    If trim(UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)) <> "" Then direct_cs_inc_det = direct_cs_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)

                    Call write_variable_with_indent_in_CASE_NOTE(direct_cs_inc_det)
                End if
            End If
        Next
    End If
    Call write_bullet_and_variable_in_CASE_NOTE("Other Child Support Income", notes_on_cses)

    'UNEA
    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
        If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
            rsdi_income_det = ""
            If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & "RSDI: $" & UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)
            If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)

            ssi_income_det = ""
            If trim(UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & "SSI: $" & UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)
            If trim(UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)

            Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " SSA income:")
            If rsdi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(rsdi_income_det)
            If ssi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(ssi_income_det)
        End If
    Next
    Call write_bullet_and_variable_in_CASE_NOTE("Other SSA Income", notes_on_ssa_income)
    Call write_bullet_and_variable_in_CASE_NOTE("VA Income", notes_on_VA_income)
    Call write_bullet_and_variable_in_CASE_NOTE("Workers Comp Income", notes_on_WC_income)

    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
        If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
            uc_income_det_one = ""
            uc_income_det_two = ""
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) <> "" Then
                uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
                uc_income_det_one = uc_income_det_one & "UC weekly gross income: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ". "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
            Else
                uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
                If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
            End If
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Current UC account balance: $" & UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Retro Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Prosp Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC SNAP budgeted Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb) & ". "

            Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " Unemployment Income:")
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("UC Income started on: " & UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb) & ". ")
            If uc_income_det_one <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_one)
            If uc_income_det_two <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_two)
            If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = TRUE Then Call write_variable_with_indent_in_CASE_NOTE("TIKL set to check for end of UC on: " & UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb))
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Notes: " & UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb))
        End If
    Next
    Call write_bullet_and_variable_in_CASE_NOTE("Other UC Income", other_uc_income_notes)
    Call write_bullet_and_variable_in_CASE_NOTE("Unearned Income", notes_on_other_UNEA)

    If case_has_personal = TRUE Then
        If trim(cit_id) <> "" Then case_has_hc_personal = TRUE
        If trim(IMIG) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_acci_info) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_faci_info) <> "" Then case_has_hc_personal = TRUE
        If trim(waiver_ltc_info) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_medi_info) <> "" Then case_has_hc_personal = TRUE
        If trim(hc_insa_info) <> "" Then case_has_hc_personal = TRUE
        If trim(DISA) <> "" Then case_has_hc_personal = TRUE
        If trim(PREG) <> "" Then case_has_hc_personal = TRUE
        If trim(ABPS) <> "" Then case_has_hc_personal = TRUE
        If trim(AREP) <> "" Then case_has_hc_personal = TRUE
        If trim(DISQ) <> "" Then case_has_hc_personal = TRUE
    End If
    If case_has_hc_personal = TRUE Then Call write_variable_in_CASE_NOTE("===== PERSONAL =====")

    Call write_bullet_and_variable_in_CASE_NOTE("Citizenship/ID", cit_id)
    Call write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
    Call write_bullet_and_variable_in_CASE_NOTE("Changes", case_changes)
    Call write_bullet_and_variable_in_CASE_NOTE("Accident", hc_acci_info)

    Call write_bullet_and_variable_in_CASE_NOTE("Facility", hc_faci_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Waiver/LTC", waiver_ltc_info)

    Call write_bullet_and_variable_in_CASE_NOTE("Medicare", hc_medi_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Insurance", hc_insa_info)

    Call write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
    Call write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
    Call write_bullet_and_variable_in_CASE_NOTE("Absent Parent", ABPS)
    If CS_forms_sent_date <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Child Support Forms given/sent to client on " & CS_forms_sent_date)
    Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)

    'DISQ
    Call write_bullet_and_variable_in_CASE_NOTE("DISQ", DISQ)

    If case_has_expenses = TRUE Then
        If trim(notes_on_coex) <> "" Then case_has_hc_expenses = TRUE
        If trim(notes_on_dcex) <> "" Then case_has_hc_expenses = TRUE
        If trim(notes_on_other_deduction) <> "" Then case_has_hc_expenses = TRUE
        If trim(hc_bils_info) <> "" Then case_has_hc_expenses = TRUE
        If trim(expense_notes) <> "" Then case_has_hc_expenses = TRUE
    End If
    If case_has_hc_expenses = TRUE Then
        Call write_variable_in_CASE_NOTE("===== EXPENSES =====")
    Else
        Call write_variable_in_CASE_NOTE("== No expense detail for this case ==")
    End If

    'Expenses
    Call write_bullet_and_variable_in_CASE_NOTE("Court Ordered Expenses", notes_on_coex)
    Call write_bullet_and_variable_in_CASE_NOTE("Dependent Care Expenses", notes_on_dcex)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Expenses", notes_on_other_deduction)
    Call write_bullet_and_variable_in_CASE_NOTE("Medical Bills", hc_bils_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Expense Detail", expense_notes)

    If case_has_resources = TRUE Then
        Call write_variable_in_CASE_NOTE("===== RESOURCES =====")
    Else
        Call write_variable_in_CASE_NOTE("== No resource/asset detail for this case ==")
    End If
    'Assets
    If confirm_no_account_panel_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Income sources have been reviewed for direct deposit/associated accounts and none were found.")
    Call write_bullet_and_variable_in_CASE_NOTE("Accounts", notes_on_acct)
    Call write_bullet_and_variable_in_CASE_NOTE("Cash", notes_on_cash)
    Call write_bullet_and_variable_in_CASE_NOTE("Cars", notes_on_cars)
    Call write_bullet_and_variable_in_CASE_NOTE("Real Estate", notes_on_rest)
    Call write_bullet_and_variable_in_CASE_NOTE("Other Assets", notes_on_other_assets)

    Call write_variable_in_CASE_NOTE("=====================")
    Call write_bullet_and_variable_in_CASE_NOTE("Notes", hc_other_notes)
    Call write_bullet_and_variable_in_CASE_NOTE("Verifications Needed", hc_verifs_needed)

    If MMIS_updated_check = checked Then Call write_variable_in_CASE_NOTE("* MMIS Updated")
    If hc_tikl_checkbox = checked Then Call write_variable_in_CASE_NOTE("* TIKL set for 45 days from application date.")
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
    EMReadScreen top_note_header, 55, 5, 25
    case_notes_information = case_notes_information & "MX Header - " & top_note_header & " %^% %^%"
    save_your_work

    Call back_to_SELF

End If

If the_process_for_snap = "Application" AND exp_det_case_note_found = False Then
    If full_determination_done = True Then
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
                objTextStream.WriteLine "WORKER USER ID ^*^*^" & user_ID_for_validation
        		objTextStream.WriteLine "CASE X NUMBER  ^*^*^" & case_pw
                If IsDate(CAF_datestamp) = True Then CAF_datestamp = DateAdd("d", 0, CAF_datestamp)
        		objTextStream.WriteLine "DATE OF APPLICATION ^*^*^" & CAF_datestamp
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
                objTextStream.WriteLine "SCRIPT RUN ^*^*^CAF"

        		'Close the object so it can be opened again shortly
        		objTextStream.Close

        	End With

        End if

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
            Call write_variable_in_case_note ("Info from INITIAL EXPEDTIED SCREENING (resident reported on Application)")
        	Call write_variable_in_case_note ("  Expedited Screening found: " & xfs_screening)
        	Call write_variable_in_case_note ("  Based on: Income:  $ " & right("        " & caf_one_income, 8) & ", Assets:    $ " & right("        " & caf_one_assets, 8)    & ", Totaling: $ " & right("        " & caf_one_resources, 8))
        	Call write_variable_in_case_note ("            Shelter: $ " & right("        " & caf_one_rent, 8)   & ", Utilities: $ " & right("        " & caf_one_utilities, 8) & ", Totaling: $ " & right("        " & caf_one_expenses, 8))
            Call write_variable_in_case_note ("No case action can be taken from screening alone, info may change at intrvw.")
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
            Call write_variable_in_case_note ("Info from Interview - Expedited Determination Completed:")
        	IF is_elig_XFS = TRUE Then
        		Call write_variable_in_case_note ("  Case is determined to meet criteria for Expedited SNAP.")
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
        	IF is_elig_XFS = FALSE Then Call write_variable_in_case_note ("  Case does not meet Expedited SNAP criteria.")
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

        PF3
    End If
End If

'Verification NOTE
verifs_needed = replace(verifs_needed, "[Information here creates a SEPARATE CASE/NOTE.]", "")
If trim(verifs_needed) = "" or verifications_requested_case_note_found = True Then
    case_notes_information = case_notes_information & "No Verifs NOTE Attempted "
    If trim(verifs_needed) = "" Then case_notes_information = case_notes_information & "- verif field is blank"
    If verifications_requested_case_note_found = True Then case_notes_information = case_notes_information & "- verif case note was found"
    case_notes_information = case_notes_information & " %^% %^%"
End If
If trim(verifs_needed) <> "" AND verifications_requested_case_note_found = False Then

    verif_counter = 1
    verifs_needed = trim(verifs_needed)
    If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
    If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
    If InStr(verifs_needed, ";") <> 0 Then
        verifs_array = split(verifs_needed, ";")
    Else
        verifs_array = array(verifs_needed)
    End If

    programs_verifs_apply_to = ""
    If verif_snap_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", SNAP"
    If verif_cash_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", CASH"
    If verif_mfip_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", MFIP"
    If verif_dwp_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", DWP"
    If verif_msa_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", MSA"
    If verif_ga_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", GA"
    If verif_grh_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", GRH"
    If verif_emer_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", EMER"
    If verif_hc_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", HC"
    If left(programs_verifs_apply_to, 1) = "," Then programs_verifs_apply_to = right(programs_verifs_apply_to, len(programs_verifs_apply_to)-1)
    programs_verifs_apply_to = trim(programs_verifs_apply_to)

    case_notes_information = case_notes_information & "Verifs NOTE Attempted %^%"
    case_notes_information = case_notes_information & "Script Header - " & "VERIFICATIONS REQUESTED" & " %^%"
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
    If programs_verifs_apply_to <> "" Then
        Call write_variable_in_CASE_NOTE("---")
        Call write_variable_in_CASE_NOTE("Verifications are needed for " & programs_verifs_apply_to & ".")
    End If
    If verifs_postponed_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("---")
        Call write_variable_in_CASE_NOTE("There may be verifications that are postponed to allow for the approval of Expedited SNAP.")
    End If
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
    EMReadScreen top_note_header, 55, 5, 25
    case_notes_information = case_notes_information & "MX Header - " & top_note_header & " %^% %^%"
    save_your_work

    Call back_to_SELF
End If

If qual_questions_yes = False or caf_qualifying_questions_case_note_found = True Then
    case_notes_information = case_notes_information & "No Qualifying Questions NOTE Attempted "
    If qual_questions_yes = False Then case_notes_information = case_notes_information & "- no qualifying questions were yes"
    If caf_qualifying_questions_case_note_found = True Then case_notes_information = case_notes_information & "- qualifying questions case note was found"
    case_notes_information = case_notes_information & " %^% %^%"
End If
If qual_questions_yes = TRUE AND caf_qualifying_questions_case_note_found = False Then
    case_notes_information = case_notes_information & "Qualifying Questions NOTE Attempted %^%"
    case_notes_information = case_notes_information & "Script Header - " & "CAF Qualifying Questions had an answer of 'YES' for at least one question" & " %^%"
    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("CAF Qualifying Questions had an answer of 'YES' for at least one question")
    If qual_question_one = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fraud/DISQ for IPV (program violation)", qual_memb_one)
    If qual_question_two = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("SNAP in more than One State", qual_memb_two)
    If qual_question_three = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fleeing Felon", qual_memb_three)
    If qual_question_four = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Drug Felony", qual_memb_four)
    If qual_question_five = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Parole/Probation Violation", qual_memb_five)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3
    EMReadScreen top_note_header, 55, 5, 25
    case_notes_information = case_notes_information & "MX Header - " & top_note_header & " %^% %^%"
    save_your_work

    Call back_to_SELF
End If

'MAIN CAF Information NOTE
'Navigates to case note, and checks to make sure we aren't in inquiry.
case_notes_information = case_notes_information & "MAIN CAF NOTE Attempted %^%"
Call start_a_blank_CASE_NOTE

If CAF_form = "HUF (DHS-8107)" Then
    case_notes_information = case_notes_information & "Script Header - " & CAF_datestamp & " HUF for " & prog_and_type_list & ": " & CAF_status & " %^% %^%"
    CALL write_variable_in_CASE_NOTE(CAF_datestamp & " HUF for " & prog_and_type_list & ": " & CAF_status)
Else
    case_notes_information = case_notes_information & "Script Header - " & CAF_datestamp & " CAF for " & prog_and_type_list & ": " & CAF_status & " %^% %^%"
    CALL write_variable_in_CASE_NOTE(CAF_datestamp & " CAF for " & prog_and_type_list & ": " & CAF_status)
End If
If multiple_CAF_dates = True Then
	CALL write_variable_in_CASE_NOTE("  --- Multiple forms received ---")
	Call write_bullet_and_variable_in_CASE_NOTE("Form", REVW_CAF_Form & ", received on: " & REVW_CAF_datestamp)
	Call write_bullet_and_variable_in_CASE_NOTE("Form", PROG_CAF_Form & ", received on: " & PROG_CAF_datestamp)
Else
	Call write_bullet_and_variable_in_CASE_NOTE("Form Received", CAF_form)
	Call write_bullet_and_variable_in_CASE_NOTE("Form Received Date", CAF_datestamp)
End If
'Programs requested
If CASH_checkbox = checked Then CAF_progs = CAF_progs & ", Cash"
If GRH_checkbox = checked Then CAF_progs = CAF_progs & ", GRH"
If SNAP_checkbox = checked Then CAF_progs = CAF_progs & ", SNAP"
If EMER_checkbox = checked Then CAF_progs = CAF_progs & ", EMER"
If CAF_progs <> "" Then
    CAF_progs = right(CAF_progs, len(CAF_progs) - 2)
Else
    CAF_progs = "None"
End If
Call write_bullet_and_variable_in_CASE_NOTE("Programs requested", CAF_progs)
If CASH_checkbox = checked Then
	If family_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for FAMILY programs.")
	If adult_cash = True Then Call write_variable_in_CASE_NOTE("* Cash request is for ADULT programs.")
End If
Call write_bullet_and_variable_in_CASE_NOTE("Info", case_details_and_notes_about_process)
'Household and personal information
If SNAP_checkbox = checked Then
    Call write_variable_in_CASE_NOTE("* SNAP unit consists of " & total_snap_count & " people - " & adult_snap_count & " adults and " & child_snap_count & " children.")
    If included_snap_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members on SNAP grant: " & included_snap_members)
    If counted_snap_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for SNAP: " & counted_snap_members)
    If EATS <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Information on EATS: " & EATS)
End If
If cash_checkbox = checked Then
    Call write_variable_in_CASE_NOTE("* CASH unit consists of " & total_cash_count & " people - " & adult_cash_count & " adults and " & child_cash_count & " children.")
    If pregnant_caregiver_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Pregnant Caregiver on Grant.")
    If included_cash_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members on CASH grant: " & included_cash_members)
    If counted_cash_members <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for CASH: " & counted_cash_members)
End If
If EMER_checkbox = checked Then
    Call write_variable_in_CASE_NOTE("* EMER unit consists of " & total_emer_count & " people - " & adult_emer_count & " adults and " & child_emer_count & " children.")
    Call write_variable_with_indent_in_CASE_NOTE("Members on EMER grant: " & included_emer_members)
    Call write_variable_with_indent_in_CASE_NOTE("Members with income counted ONLY for EMER: " & counted_emer_members)
End If
Call write_bullet_and_variable_in_CASE_NOTE("Relationships", relationship_detail)

first_member = TRUE
For the_member = 0 to UBound(ALL_MEMBERS_ARRAY, 2)
    If ALL_MEMBERS_ARRAY(gather_detail, the_member) = TRUE Then
        If ALL_MEMBERS_ARRAY(id_required, the_member) = checked Then
            If first_member = TRUE Then
                Call write_variable_in_CASE_NOTE("===== ID REQUIREMENT =====")
                first_member = FALSE
            End If
            Call write_variable_in_CASE_NOTE("* Identity of Memb " & ALL_MEMBERS_ARRAY(memb_numb, the_member) & " verified by: " & right(ALL_MEMBERS_ARRAY(clt_id_verif, the_member), len(ALL_MEMBERS_ARRAY(clt_id_verif, the_member)) - 5) & " and is required.")
            If trim(ALL_MEMBERS_ARRAY(id_detail, the_member)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Details: " & trim(ALL_MEMBERS_ARRAY(id_detail, the_member)))
        End If
    End If
Next

'INCOME
If case_has_income = TRUE Then
    Call write_variable_in_CASE_NOTE("===== INCOME =====")
Else
    Call write_variable_in_CASE_NOTE("== No Income detail Listed for this case. ==")
End If
'JOBS
If ALL_JOBS_PANELS_ARRAY(memb_numb, 0) <> "" Then
    ' Call write_variable_with_indent(variable_name)
    Call write_variable_in_CASE_NOTE("--- JOBS Income ---")
    For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
        Call write_variable_in_CASE_NOTE("Member " & ALL_JOBS_PANELS_ARRAY(memb_numb, each_job) & " at " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job))
        If ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = checked Then Call write_variable_in_CASE_NOTE("* This job has not been verified and this is only an estimate.")
        IF ALL_JOBS_PANELS_ARRAY(EI_case_note, each_job) = TRUE Then call write_variable_in_CASE_NOTE("* BUDGET DETAIL ABOUT THIS JOB IN PREVIOUS CASE NOTE.")
        If ALL_JOBS_PANELS_ARRAY(verif_code, each_job) = "Delayed" Then
            Call write_variable_in_CASE_NOTE("* Verification of this job has been delayed for review or approval of Expedited SNAP.")
        ElseIf ALL_JOBS_PANELS_ARRAY(estimate_only, each_job) = unchecked Then
            Call write_variable_in_CASE_NOTE("* Verification - " & ALL_JOBS_PANELS_ARRAY(verif_code, each_job))
        End If
        Call write_bullet_and_variable_in_CASE_NOTE("Verification", ALL_JOBS_PANELS_ARRAY(verif_explain, each_job))
        If ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Retro Income: $" & ALL_JOBS_PANELS_ARRAY(job_retro_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(retro_hours, each_job) & " hours.")
        If ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Prospective Income: $" & ALL_JOBS_PANELS_ARRAY(job_prosp_income, each_job) & " - " & ALL_JOBS_PANELS_ARRAY(prosp_hours, each_job) & " hours.")
        If snap_checkbox = checked Then Call write_variable_with_indent_in_CASE_NOTE("SNAP Budget Detail: Monthly budgeted amount - $" & ALL_JOBS_PANELS_ARRAY(pic_prosp_income, each_job) & " based on $" & ALL_JOBS_PANELS_ARRAY(pic_pay_date_income, each_job) & " paid " & ALL_JOBS_PANELS_ARRAY(pic_pay_freq, each_job) & ". Calculated on " & ALL_JOBS_PANELS_ARRAY(pic_calc_date, each_job))
        If ALL_JOBS_PANELS_ARRAY(budget_explain, each_job) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("About Budget: " & ALL_JOBS_PANELS_ARRAY(budget_explain, each_job))
    Next
End If
Call write_bullet_and_variable_in_CASE_NOTE("JOBS", notes_on_jobs)
Call write_bullet_and_variable_in_CASE_NOTE("Other Earned Income", earned_income)

'BUSI
If ALL_BUSI_PANELS_ARRAY(memb_numb, 0) <> "" Then
    Call write_variable_in_CASE_NOTE("--- BUSI Income ---")
    For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
        busi_det_msg = "Self Employment for Memb " & ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) & " - BUSI type:" & right(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi), len(ALL_BUSI_PANELS_ARRAY(busi_type, each_busi)) - 4) & "."
        Call write_variable_in_CASE_NOTE(busi_det_msg)

        If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Description: " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) & ".")
        If ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business structure: " & ALL_BUSI_PANELS_ARRAY(busi_structure, each_busi) & ".")
        If ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Clt owns " & ALL_BUSI_PANELS_ARRAY(share_num, each_busi) & "/" & ALL_BUSI_PANELS_ARRAY(share_denom, each_busi) & " of the business.")
        If ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Business also owned by Memb(s) " & ALL_BUSI_PANELS_ARRAY(partners_in_HH, each_busi) & ".")

        se_method_det_msg = "* Self Employment Budgeting method selected: " & ALL_BUSI_PANELS_ARRAY(calc_method, each_busi) & "."
        Call write_variable_in_CASE_NOTE(se_method_det_msg)
        If ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Method selected on: " & ALL_BUSI_PANELS_ARRAY(mthd_date, each_busi) & ".")
        If ALL_BUSI_PANELS_ARRAY(method_convo_checkbox, each_busi) = checked Then Call write_variable_with_indent_in_CASE_NOTE("The self employment method selected was discussed with the client.")

        If cash_checkbox = checked OR EMER_checkbox = checked Then
            Call write_variable_in_CASE_NOTE("* Cash Income and Expense Detail:")
            cash_income_det = ""
            cash_expense_det = ""

            If ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) <> "" Then
                cash_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_cash, each_busi) & " - "
                cash_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_cash, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) <> "" Then
                cash_income_det = cash_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_cash, each_busi) & " - "
                cash_expense_det = cash_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_cash, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi) <> "" Then cash_income_det = cash_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_income_verif, each_busi)
            If ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi) <> "" Then cash_expense_det = cash_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(cash_expense_verif, each_busi)

            Call write_variable_with_indent_in_CASE_NOTE(cash_income_det)
            Call write_variable_with_indent_in_CASE_NOTE(cash_expense_det)
        End If
        If SNAP_checkbox = checked Then
            Call write_variable_in_CASE_NOTE("* SNAP Income and Expense Detail:")
            snap_income_det = ""
            snap_expense_det = ""

            If ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) <> "" Then
                snap_income_det = "RETRO Income $" & ALL_BUSI_PANELS_ARRAY(income_ret_snap, each_busi) & " - "
                snap_expense_det = "RETRO Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_ret_snap, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) <> "" Then
                snap_income_det = snap_income_det & "PROSP Income $" & ALL_BUSI_PANELS_ARRAY(income_pro_snap, each_busi) & " - "
                snap_expense_det = snap_expense_det & "PROSP Expenses $" & ALL_BUSI_PANELS_ARRAY(expense_pro_snap, each_busi) & " - "
            End If
            If ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi) <> "" Then snap_income_det = snap_income_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_income_verif, each_busi)
            If ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "Select or Type" and ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi) <> "" Then snap_expense_det = snap_expense_det & "Verification: " & ALL_BUSI_PANELS_ARRAY(snap_expense_verif, each_busi)

            Call write_variable_with_indent_in_CASE_NOTE(snap_income_det)
            Call write_variable_with_indent_in_CASE_NOTE(snap_expense_det)
            If ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Expenses from taxes not allowed: " & ALL_BUSI_PANELS_ARRAY(exp_not_allwd, each_busi))
        End If
        rept_hours_det_msg = ""
        min_wg_hours_det_msg = ""
        If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) = ""
        If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) = ""
        If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) = ""
        If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = "___" Then ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) = ""

        If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then
            rept_hours_det_msg = rept_hours_det_msg & "Clt reported monthly work hours of: "
            If ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_retro_hrs, each_busi) & " retrospecive work and "
            If ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) <> "" Then rept_hours_det_msg = rept_hours_det_msg & ALL_BUSI_PANELS_ARRAY(rept_prosp_hrs, each_busi) & " prospoective work hrs"
            rept_hours_det_msg = rept_hours_det_msg & ". "
        End If
        If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" OR ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then
            min_wg_hours_det_msg = min_wg_hours_det_msg & "Work earnings indicate Minumun Wage Hours of: "
            If ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_retro_hrs, each_busi) & " retrospective and "
            If ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) <> "" Then min_wg_hours_det_msg = min_wg_hours_det_msg & ALL_BUSI_PANELS_ARRAY(min_wg_prosp_hrs, each_busi) & " prospective"
            min_wg_hours_det_msg = min_wg_hours_det_msg & ". "
        End If
        If rept_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & rept_hours_det_msg)
        If min_wg_hours_det_msg <> "" Then Call write_variable_in_CASE_NOTE("* " & min_wg_hours_det_msg)
        If ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Verif Detail: " & ALL_BUSI_PANELS_ARRAY(verif_explain, each_busi))
        If ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Budget Detail: " & ALL_BUSI_PANELS_ARRAY(budget_explain, each_busi))
    Next
End If
Call write_bullet_and_variable_in_CASE_NOTE("BUSI", notes_on_busi)

'CSES
If show_cses_detail = TRUE Then
    Call write_variable_in_CASE_NOTE("--- Child Support Income ---")
    For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
        If UNEA_INCOME_ARRAY(CS_exists, each_unea_memb) = TRUE Then
            total_cs = 0
            If IsNumeric(UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
            If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
            If IsNumeric(UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)) = TRUE Then total_cs = total_cs + UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)

            Call write_variable_in_CASE_NOTE("* Total child support income for Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & ": $" & total_cs)
            If UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb) <> "" Then
                cs_disb_inc_det = "Disbursed child support: $" & UNEA_INCOME_ARRAY(disb_CS_amt, each_unea_memb)
                If trim(UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)) <> "" Then cs_disb_inc_det = cs_disb_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_notes, each_unea_memb)

                Call write_variable_with_indent_in_CASE_NOTE(cs_disb_inc_det)
                If trim(UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_months, each_unea_memb) & " month(s) of disbursement income.")
                If trim(UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_prosp_budg, each_unea_memb))
            End If

            If UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb) <> "" Then
                cs_arrears_inc_det = "Disbursed child support arrears: $" & UNEA_INCOME_ARRAY(disb_CS_arrears_amt, each_unea_memb)
                If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)) <> "" Then cs_arrears_inc_det = cs_arrears_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(disb_CS_arrears_notes, each_unea_memb)

                Call write_variable_with_indent_in_CASE_NOTE(cs_arrears_inc_det)
                If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Income was determined using " & UNEA_INCOME_ARRAY(disb_CS_arrears_months, each_unea_memb) & " month(s) of disbursement income.")
                If trim(UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("SNAP prospective budget details: " & UNEA_INCOME_ARRAY(disb_CS_arrears_budg, each_unea_memb))
            End If

            If UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb) <> "" Then
                direct_cs_inc_det = "Direct child support: $" & UNEA_INCOME_ARRAY(direct_CS_amt, each_unea_memb)
                If trim(UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)) <> "" Then direct_cs_inc_det = direct_cs_inc_det & ". Notes: " & UNEA_INCOME_ARRAY(direct_CS_notes, each_unea_memb)

                Call write_variable_with_indent_in_CASE_NOTE(direct_cs_inc_det)
            End if
        End If
    Next
End If
Call write_bullet_and_variable_in_CASE_NOTE("Other Child Support Income", notes_on_cses)

'UNEA
For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(SSA_exists, each_unea_memb) = TRUE Then
        rsdi_income_det = ""
        If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & "RSDI: $" & UNEA_INCOME_ARRAY(UNEA_RSDI_amt, each_unea_memb)
        If trim(UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)) <> "" Then rsdi_income_det = rsdi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_RSDI_notes, each_unea_memb)

        ssi_income_det = ""
        If trim(UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & "SSI: $" & UNEA_INCOME_ARRAY(UNEA_SSI_amt, each_unea_memb)
        If trim(UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)) <> "" Then ssi_income_det = ssi_income_det & ". Notes: " & UNEA_INCOME_ARRAY(UNEA_SSI_notes, each_unea_memb)

        Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " SSA income:")
        If rsdi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(rsdi_income_det)
        If ssi_income_det <> "" Then Call write_variable_with_indent_in_CASE_NOTE(ssi_income_det)
    End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Other SSA Income", notes_on_ssa_income)
Call write_bullet_and_variable_in_CASE_NOTE("VA Income", notes_on_VA_income)
Call write_bullet_and_variable_in_CASE_NOTE("Workers Comp Income", notes_on_WC_income)

For each_unea_memb = 0 to UBound(UNEA_INCOME_ARRAY, 2)
    If UNEA_INCOME_ARRAY(UC_exists, each_unea_memb) = TRUE Then
        uc_income_det_one = ""
        uc_income_det_two = ""
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb)) <> "" Then
            uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
            uc_income_det_one = uc_income_det_one & "UC weekly gross income: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_gross, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
        Else
            uc_income_det_one = uc_income_det_one & "Budgeted UC weekly amount: $" & UNEA_INCOME_ARRAY(UNEA_UC_weekly_net, each_unea_memb) & ".; "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount allowed: $" & UNEA_INCOME_ARRAY(UNEA_UC_counted_ded, each_unea_memb) & ". "
            If trim(UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Deduction amount excluded: $" & UNEA_INCOME_ARRAY(UNEA_UC_exclude_ded, each_unea_memb) & ". "
        End If
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb)) <> "" Then uc_income_det_one = uc_income_det_one & "Current UC account balance: $" & UNEA_INCOME_ARRAY(UNEA_UC_account_balance, each_unea_memb) & ". "
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Retro Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_retro_amt, each_unea_memb) & ". "
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC Prosp Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_prosp_amt, each_unea_memb) & ". "
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb)) <> "" Then uc_income_det_two = uc_income_det_two & "UC SNAP budgeted Income: $" & UNEA_INCOME_ARRAY(UNEA_UC_monthly_snap, each_unea_memb) & ". "

        Call write_variable_in_CASE_NOTE("* Member " & UNEA_INCOME_ARRAY(memb_numb, each_unea_memb) & " Unemployment Income:")
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("UC Income started on: " & UNEA_INCOME_ARRAY(UNEA_UC_start_date, each_unea_memb) & ". ")
        If uc_income_det_one <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_one)
        If uc_income_det_two <> "" Then Call write_variable_with_indent_in_CASE_NOTE(uc_income_det_two)
        If IsDate(UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb)) = TRUE Then Call write_variable_with_indent_in_CASE_NOTE("TIKL set to check for end of UC on: " & UNEA_INCOME_ARRAY(UNEA_UC_tikl_date, each_unea_memb))
        If trim(UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb)) <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Notes: " & UNEA_INCOME_ARRAY(UNEA_UC_notes, each_unea_memb))
    End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Other UC Income", other_uc_income_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Unearned Income", notes_on_other_UNEA)

If case_has_personal = TRUE Then Call write_variable_in_CASE_NOTE("===== PERSONAL =====")

Call write_bullet_and_variable_in_CASE_NOTE("Citizenship/ID", cit_id)
Call write_bullet_and_variable_in_CASE_NOTE("IMIG", IMIG)
Call write_bullet_and_variable_in_CASE_NOTE("School", SCHL)
Call write_bullet_and_variable_in_CASE_NOTE("Changes", case_changes)
Call write_bullet_and_variable_in_CASE_NOTE("DISA", DISA)
Call write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
Call write_bullet_and_variable_in_CASE_NOTE("PREG", PREG)
Call write_bullet_and_variable_in_CASE_NOTE("Absent Parent", ABPS)
If CS_forms_sent_date <> "" Then Call write_variable_with_indent_in_CASE_NOTE("Child Support Forms given/sent to client on " & CS_forms_sent_date)
Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)

'Address Detail
If address_confirmation_checkbox = checked Then Call write_variable_in_CASE_NOTE("* The address on ADDR was reviewed and is correct.")
If homeless_yn = "Yes" Then Call write_variable_in_CASE_NOTE("* Household is homeless.")
Call write_variable_in_CASE_NOTE("* Client reports living in county " & addr_county)
Call write_bullet_and_variable_in_CASE_NOTE("Living Situation", living_situation)
Call write_bullet_and_variable_in_CASE_NOTE("Address Detail", notes_on_address)

'DISQ
Call write_bullet_and_variable_in_CASE_NOTE("DISQ", DISQ)

'WREG and ABAWD
Call write_bullet_and_variable_in_CASE_NOTE("WREG", notes_on_wreg)
all_abawd_notes = notes_on_abawd & notes_on_abawd_two & notes_on_abawd_three
Call write_bullet_and_variable_in_CASE_NOTE("ABAWD", all_abawd_notes)

Call write_bullet_and_variable_in_CASE_NOTE("Medicare", MEDI)
Call write_bullet_and_variable_in_CASE_NOTE("Diet", DIET)

'MFIP-DWP information
Call write_bullet_and_variable_in_CASE_NOTE("Time Tracking (MFIP)", notes_on_time)
Call write_bullet_and_variable_in_CASE_NOTE("MFIP Sanction", notes_on_sanction)
Call write_bullet_and_variable_in_CASE_NOTE("MF/DWP Employment Services", EMPS)

If case_has_expenses = TRUE Then
    Call write_variable_in_CASE_NOTE("===== EXPENSES =====")
Else
    Call write_variable_in_CASE_NOTE("== No expense detail for this case ==")
End If
'SHEL
Call write_bullet_and_variable_in_CASE_NOTE("Shelter Expense", "$" & total_shelter_amount)

If InStr(full_shelter_details, "*") <> 0 Then
    shelter_detail_array = split(full_shelter_details, "*")
Else
    shelter_detail_array = array(full_shelter_details)
End If
If full_shelter_details <> "" Then
    For each shel_info in shelter_detail_array
        shel_info = trim(shel_info)
        Call write_variable_with_indent_in_CASE_NOTE(shel_info)
    Next
End If
'HEST/ACUT
Call write_bullet_and_variable_in_CASE_NOTE("Actual Utility Expenses", notes_on_acut)
If hest_information <> "Select ALLOWED HEST" Then Call write_variable_in_CASE_NOTE("* Standard Utility expenses: " & hest_information)

'Expenses
Call write_bullet_and_variable_in_CASE_NOTE("Court Ordered Expenses", notes_on_coex)
Call write_bullet_and_variable_in_CASE_NOTE("Dependent Care Expenses", notes_on_dcex)
Call write_bullet_and_variable_in_CASE_NOTE("Other Expenses", notes_on_other_deduction)
Call write_bullet_and_variable_in_CASE_NOTE("Expense Detail", expense_notes)
Call write_bullet_and_variable_in_CASE_NOTE("FS Medical Expenses", FMED)

If case_has_resources = TRUE Then
    Call write_variable_in_CASE_NOTE("===== RESOURCES =====")
Else
    Call write_variable_in_CASE_NOTE("== No resource/asset detail for this case ==")
End If
'Assets
If confirm_no_account_panel_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Income sources have been reviewed for direct deposit/associated accounts and none were found.")
Call write_bullet_and_variable_in_CASE_NOTE("Accounts", notes_on_acct)
Call write_bullet_and_variable_in_CASE_NOTE("Cash", notes_on_cash)
Call write_bullet_and_variable_in_CASE_NOTE("Cars", notes_on_cars)
Call write_bullet_and_variable_in_CASE_NOTE("Real Estate", notes_on_rest)
Call write_bullet_and_variable_in_CASE_NOTE("Other Assets", notes_on_other_assets)

Call write_variable_in_CASE_NOTE("===== Case Information =====")
'Next review
If trim(next_er_month) <> "" Then Call write_bullet_and_variable_in_CASE_NOTE("Next ER", next_er_month & "/" & next_er_year)

IF application_signed_checkbox = checked THEN
	CALL write_variable_in_CASE_NOTE("* Application was signed.")
Else
	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
END IF
IF eDRS_sent_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* eDRS sent.")
IF updated_MMIS_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Updated MMIS.")
IF WF1_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Workforce referral made.")

IF Sent_arep_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Sent form(s) to AREP.")
IF intake_packet_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Client received intake packet.")
IF IAA_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* IAAs/OMB given to client.")

IF client_delay_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
If TIKL_checkbox Then CALL write_variable_in_CASE_NOTE("* TIKL set to take action on " & DateAdd("d", 30, CAF_datestamp))
If client_delay_TIKL_checkbox Then CALL write_variable_in_CASE_NOTE("* TIKL set to update PND2 for Client Delay on " & DateAdd("d", 10, CAF_datestamp))

If qual_questions_yes = FALSE Then Call write_variable_in_CASE_NOTE("* All Qualifying Questions answered 'No'.")
Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
If trim(verifs_needed) <> "" Then Call write_variable_in_CASE_NOTE("** VERIFICATIONS REQUESTED - See previous case note for detail")
' IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF SNAP_recert_is_likely_24_months = TRUE THEN					'if we determined on stat/revw that the next SNAP recert date isn't 12 months beyond the entered footer month/year
	TIKL_for_24_month = msgbox("Your SNAP recertification date is listed as " & SNAP_recert_date & " on STAT/REVW. Do you want set a TIKL on " & dateadd("m", "-1", SNAP_recert_compare_date) & " for 12 month contact?" & vbCR & vbCR & "NOTE: Clicking yes will navigate away from CASE/NOTE saving your case note.", VBYesNo)
	IF TIKL_for_24_month = vbYes THEN 												'if the select YES then we TIKL using our custom functions.
		'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
        Call create_TIKL("If SNAP is open, review to see if 12 month contact letter is needed. DAIL scrubber can send 12 Month Contact Letter if used on this TIKL.", 0, dateadd("m", "-1", SNAP_recert_compare_date), False, TIKL_note_text)
        Call back_to_SELF
	END IF
END IF

end_msg = "Success! " & CAF_form & " has been successfully noted. Please remember to run the Eligibility Summary script if results have been approved in MAXIS ('APP' completed)."

save_your_work

revw_pending_table = False                                                      'Determining if we should be adding this case to the CasesPending SQL Table
If unknown_cash_pending = True Then revw_pending_table = True                   'case should be pending cash or snap and NOT have SNAP active
If ga_status = "PENDING" Then revw_pending_table = True
If msa_status = "PENDING" Then revw_pending_table = True
If mfip_status = "PENDING" Then revw_pending_table = True
If dwp_status = "PENDING" Then revw_pending_table = True
If grh_status = "PENDING" Then revw_pending_table = True
If snap_status = "PENDING" Then revw_pending_table = True
If snap_status = "ACTIVE" Then revw_pending_table = False

'Here we go to ensure this case is listed in the CasesPending table for ES Workflow
If developer_mode = False AND revw_pending_table = True Then                    'Only do this if not in training region.
	MAXIS_case_number = trim(MAXIS_case_number)
    eight_digit_case_number = right("00000000"&MAXIS_case_number, 8)            'The SQL table functionality needs the leading 0s added to the Case Number

    If unknown_cash_pending = True Then cash_stat_code = "P"                    'determining the program codes for the table entry

    If ma_status = "INACTIVE" Or ma_status = "APP CLOSE" Then hc_stat_code = "I"
    If ma_status = "ACTIVE" Or ma_status = "APP OPEN" Then hc_stat_code = "A"
    If ma_status = "REIN" Then hc_stat_code = "R"
    If ma_status = "PENDING" Then hc_stat_code = "P"
    If msp_status = "INACTIVE" Or msp_status = "APP CLOSE" Then hc_stat_code = "I"
    If msp_status = "ACTIVE" Or msp_status = "APP OPEN" Then hc_stat_code = "A"
    If msp_status = "REIN" Then hc_stat_code = "R"
    If msp_status = "PENDING" Then hc_stat_code = "P"
    If unknown_hc_pending = True Then hc_stat_code = "P"

    If ga_status = "PENDING" Then ga_stat_code = "P"
    If ga_status = "REIN" Then ga_stat_code = "R"
    If ga_status = "ACTIVE" Or ga_status = "APP OPEN" Then ga_stat_code = "A"
    If ga_status = "INACTIVE" Or ga_status = "APP CLOSE" Then ga_stat_code = "I"

    If grh_status = "PENDING" Then grh_stat_code = "P"
    If grh_status = "REIN" Then grh_stat_code = "R"
    If grh_status = "ACTIVE" Or grh_status = "APP OPEN" Then grh_stat_code = "A"
    If grh_status = "INACTIVE" Or grh_status = "APP CLOSE" Then grh_stat_code = "I"

    If emer_status = "PENDING" Then emer_stat_code = "P"
    If emer_status = "REIN" Then emer_stat_code = "R"
    If emer_status = "ACTIVE" Or emer_status = "APP OPEN" Then emer_stat_code = "A"
    If emer_status = "INACTIVE" Or emer_status = "APP CLOSE" Then emer_stat_code = "I"

    If mfip_status = "PENDING" Then mfip_stat_code = "P"
    If mfip_status = "REIN" Then mfip_stat_code = "R"
    If mfip_status = "ACTIVE" Or mfip_status = "APP OPEN" Then mfip_stat_code = "A"
    If mfip_status = "INACTIVE" Or mfip_status = "APP CLOSE" Then mfip_stat_code = "I"

    If snap_status = "PENDING" Then snap_stat_code = "P"
    If snap_status = "REIN" Then snap_stat_code = "R"
    If snap_status = "ACTIVE" Or snap_status = "APP OPEN" Then snap_stat_code = "A"
    If snap_status = "INACTIVE" Or snap_status = "APP CLOSE" Then snap_stat_code = "I"

    appears_expedited_for_data_table = 1                                        'Setting if case is Expedited or not based on information in the Determination.
    If is_elig_XFS = False Then appears_expedited_for_data_table = 0

    If IsDate(CAF_datestamp) = True Then CAF_datestamp = DateAdd("d", 0, CAF_datestamp)     'make sure that CAF date is formatted as a date

    'Setting constants
    Const adOpenStatic = 3
    Const adLockOptimistic = 3

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the BZST connection to SQL Database'
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

    'delete a record if the case number matches
    objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection

    'if one was found we are going to delete that record
    If current_case_record_found = True Then objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection

    'Add a new record with this case information'
    objRecordSet.Open "INSERT INTO ES.ES_CasesPending (WorkerID, CaseNumber, CaseName, ApplDate, FSStatusCode, CashStatusCode, HCStatusCode, GAStatusCode, GRStatusCode, EAStatusCode, MFStatusCode, IsExpSnap, UpdateDate)" &  _
                      "VALUES ('" & worker_id_for_data_table & "', '" & eight_digit_case_number & "', '" & case_name_for_data_table & "', '" & CAF_datestamp & "', '" & snap_stat_code & "', '" & cash_stat_code & "', '" & hc_stat_code & "', '" & ga_stat_code & "', '" & grh_stat_code & "', '" & emer_stat_code & "', '" & mfip_stat_code & "', '" & appears_expedited_for_data_table & "', '" & date & "')", objConnection, adOpenStatic, adLockOptimistic

    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
End If

script_end_procedure_with_error_report(end_msg)
