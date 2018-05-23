'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PAYSTUBS RECEIVED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 473                	'manual run time in seconds
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("04/23/2018", "Fixed bug in which the lines of the PIC were dupicated in the case note.", "Casey Love, Hennepin County")
CALL changelog_update("12/07/2017", "Removed condition to allow paystubs dated with the current date to be accepted. Updated code to write JOBS verification code in.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/11/2017", "The script has been updated to write to the GRH PIC and to case note that the GRH PIC has been updated.", "Robert Fewins-Kalb, Anoka County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Find case number and footer month - set a variable with the initially found footer month for the default for every loop
'The footer month may be different for EVERY income source. NEED to add handling to identify if there is a begin date for updating MAXIS (app date that activates the case)
    'A client may apply in april and bring in checks from March but we cannot update MAXIS in March


'DIALOG TO GET CASE NUMBER
'Possibly add worker signature here and take it out of the following dialogs
BeginDialog Dialog1, 0, 0, 191, 135, "Case Number"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 90, 5, 70, 15, MAXIS_case_number
  Text 5, 25, 65, 10, "Worker Signature:"
  EditBox 5, 35, 175, 15, worker_signature
  GroupBox 5, 55, 175, 55, "INSTRUCTIONS - PLEASE READ!!!"
  Text 15, 70, 155, 30, "This script will allow you to update any JOBS/BUSI/RBIC on a case. It can process multiple panels in one run. "
  ButtonGroup ButtonPressed
    OkButton 85, 115, 50, 15
    CancelButton 140, 115, 50, 15
EndDialog


'CREATE ARRAY OF ALL EI panels'
'Put them in a 'FOR-NEXT' to loop through each panel.
'IF all income will be case noted as 1 note then create an ARRAY of all the case note information.


'NAVIGATE TO JOBS for each HH MEMBER and ask if Income information was received for this job.

'This will become dynamic and there will be an array of all the checks listed.
'STILL need some handling for scheduled income with no actual checks or cases where scheduled income is different from actual checks but we get both.
'NEED TO ADD CHECKBOXES FOR PROGRAMS THIS INCOME APPLIES TO - and precheck all the programs that are active on this case'
BeginDialog Dialog1, 0, 0, 606, 160, "Enter ALL Paychecks Received"
  Text 10, 10, 265, 10, "JOBS 01 01 - EMPLOYER"
  Text 310, 10, 50, 10, "Income Type:"
  DropListBox 365, 10, 100, 45, "J - WIOA"+chr(9)+"W - Wages (Incl Tips)"+chr(9)+"E - EITC"+chr(9)+"G - Experience Works"+chr(9)+"F - Federal Work Study"+chr(9)+"S - State Work Study"+chr(9)+"O - Other"+chr(9)+"C - Contract Income"+chr(9)+"T - Training Program"+chr(9)+"P - Service Program"+chr(9)+"R - Rehab Program", income_type
  GroupBox 475, 5, 125, 25, "Apply Income to Programs:"
  CheckBox 485, 15, 30, 10, "SNAP", apply_to_SNAP
  CheckBox 530, 15, 30, 10, "CASH", apply_to_CASH
  CheckBox 570, 15, 20, 10, "HC", apply_to_HC
  Text 5, 40, 60, 10, "JOBS Verif Code:"
  DropListBox 65, 35, 105, 45, "1 - Pay Stubs/Tip Report"+chr(9)+"2 - Empl Statement"+chr(9)+"3 - Coltrl Stmt"+chr(9)+"4 - Other Document"+chr(9)+"5 - Pend Out State Verification"+chr(9)+"N - No Ver Prvd", JOBS_verif_code
  Text 175, 40, 155, 10, "additional detail of verification received:"
  EditBox 310, 35, 290, 15, Edit2
  Text 5, 60, 90, 10, "Date verification received:"
  EditBox 100, 55, 50, 15, verif_date
  Text 5, 80, 80, 10, "Pay Date (MM/DD/YY):"
  Text 90, 80, 50, 10, "Gross Amount:"
  Text 145, 80, 25, 10, "Hours:"
  Text 180, 65, 25, 25, "Use in SNAP budget"
  Text 235, 80, 85, 10, "If not used, explain why:"
  Text 355, 70, 245, 10, "If there is a specific amount that should be NOT budgeted from this check:"
  Text 355, 80, 30, 10, "Amount:"
  Text 415, 80, 30, 10, "Reason:"
  EditBox 5, 90, 65, 15, pay_date
  EditBox 90, 90, 45, 15, gross_amount
  EditBox 145, 90, 25, 15, hours_on_check
  OptionGroup RadioGroup1
    RadioButton 180, 90, 25, 10, "Yes", budget_yes
    RadioButton 210, 90, 25, 10, "No", budget_no
  EditBox 235, 90, 115, 15, reason_not_budgeted
  EditBox 355, 90, 45, 15, not_budgeted_amount
  EditBox 410, 90, 185, 15, amount_not_budgeted_reason
  Text 5, 115, 70, 10, "Anticipated Income"
  Text 5, 130, 50, 10, "Rate of Pay/Hr"
  Text 75, 130, 35, 10, "Hours/Wk"
  Text 130, 130, 50, 10, "Pay Frequency"
  Text 225, 115, 70, 10, "Regular Non-Monthly"
  Text 225, 130, 25, 10, "Amount"
  Text 280, 130, 50, 10, "Nbr of Months"
  EditBox 5, 140, 50, 15, rate_of_pay
  EditBox 75, 140, 40, 15, hours_per_week
  DropListBox 130, 140, 85, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week", pay_frequency
  EditBox 225, 140, 40, 15, non_monthly_amt
  EditBox 280, 140, 30, 15, number_non_reg_months
  ButtonGroup ButtonPressed
    PushButton 440, 140, 15, 15, "+", add_another_check
    PushButton 460, 140, 15, 15, "-", take_a_check_away
    OkButton 495, 140, 50, 15
    CancelButton 550, 140, 50, 15
EndDialog

'Script will determine pay frequency and potentially 1st check (if not listed on JOBS)
'Script will determine the initial footer month to change by the pay dates listed.
'Script will create a budget based on the program this income applies to
'Dialog the budget and have the worker confirm - if they decline - pull the check list dialog back up and have them adjust it there.
BeginDialog Dialog1, 0, 0, 421, 240, "Confirm JOBS Budget"
  Text 10, 10, 175, 10, "JOBS 01 01 - EMPLOYER"
  Text 245, 10, 50, 10, "Pay Frequency"
  DropListBox 305, 5, 95, 45, "1 - One Time Per Month"+chr(9)+"2 - Two Times Per Month"+chr(9)+"3 - Every Other Week"+chr(9)+"4 - Every Week"+chr(9)+"5 - Other", pay_frequency
  Text 240, 30, 60, 10, "Income Start Date:"
  EditBox 305, 25, 70, 15, income_start_date
  GroupBox 5, 40, 410, 105, "SNAP Budget"
  Text 10, 50, 100, 10, "Paychecks Inclued in Budget:"
  Text 20, 65, 90, 10, "01/01/2018 - $400 - 40 hrs"
  Text 20, 75, 90, 10, "01/15/2018- $400 - 40 hrs"
  Text 10, 95, 130, 10, "Paychecks not included: 12/24/2018"
  Text 185, 50, 90, 10, "Average hourly rate of pay:"
  Text 185, 65, 90, 10, "Average weekly hours:"
  Text 185, 80, 90, 10, "Average paycheck amount:"
  Text 185, 95, 90, 10, "Monthly Budgeted Income:"
  CheckBox 10, 110, 330, 10, "Check here if you confirm that this budget is correct and is the best estimate of anticipated income.", confirm_budget_checkbox
  Text 10, 130, 60, 10, "Conversation with:"
  ComboBox 75, 125, 60, 45, " "+chr(9)+"Client - not employee"+chr(9)+"Employee"+chr(9)+"Employer", converstion_with
  Text 140, 130, 25, 10, "clarifies"
  EditBox 170, 125, 235, 15, conversation_detail
  GroupBox 5, 150, 410, 60, "CASH Budget"
  Text 15, 165, 110, 10, "Actual Paychecks to add to JOBS:"
  Text 25, 180, 90, 10, "01/01/2018 - $400 - 40 hrs"
  Text 25, 190, 90, 10, "01/15/2018- $400 - 40 hrs"
  ButtonGroup ButtonPressed
    OkButton 315, 220, 50, 15
    CancelButton 370, 220, 50, 15
EndDialog

'Worker must confirm the frequency, first pay, and footer month
'Worker will inicate if future months should be updated - default this to 'yes' as script will update retro and prospective specific to each month
'SNAP PIC, GRH PIC, HC EI EST will be checked to be updated IF any of these programs are open on the case.

'NEED to add handling for future/current changes - start or stop work - get policy on this from SNAP refresher - talk to Melissa.

'NAVIGATE to BUSI for each HH MEMBER and ask if Income Information was received for this Self Employment.

BeginDialog Dialog1, 0, 0, 486, 175, "Enter Self Employment Information"
  Text 10, 10, 180, 10, "BUSI 01 01 - CLIENT NAME"
  Text 200, 10, 80, 10, "Self Employment Type:"
  DropListBox 280, 5, 125, 45, "01 - Farming"+chr(9)+"02 - Real Estate"+chr(9)+"03 - Home Product Sales"+chr(9)+"04 - Other Sales"+chr(9)+"05 - Personal Services"+chr(9)+"06 - Paper Route"+chr(9)+"07 - In Home Daycare"+chr(9)+"08 - Rental Income"+chr(9)+"09 - Other", busi_type
  Text 10, 30, 65, 10, "Verification srouce:"
  DropListBox 90, 25, 75, 45, "1 - Income Tax Returns"+chr(9)+"2 - Receipts of Sales/Purch"+chr(9)+"3 - Client Busi Records/Ledger"+chr(9)+"6 - Other Document"+chr(9)+"N - No Ver Prvd", busi_verif_code
  Text 180, 30, 100, 10, "Amount of Income Information:"
  DropListBox 290, 25, 80, 45, "Select One..."+chr(9)+"A Full Year Totaled"+chr(9)+"Month by Month", amount_income
  Text 10, 50, 120, 10, "Self Employment Budgeting Method"
  DropListBox 135, 45, 85, 45, "01 - 50% Grosss Inc"+chr(9)+"02 - Tax Forms", busi_method
  Text 225, 50, 50, 10, "Selection Date:"
  EditBox 280, 45, 50, 15, method_selection_date
  CheckBox 30, 65, 210, 10, "Check here to confirm this method was discussed with Client.", convo_checkbox
  GroupBox 415, 5, 65, 70, "Apply Income To"
  CheckBox 425, 20, 35, 10, "SNAP", apply_to_SNAP
  CheckBox 425, 35, 35, 10, "CASH", apply_to_CASH
  CheckBox 425, 50, 25, 10, "HC", apply_to_HC
  ButtonGroup ButtonPressed
    PushButton 355, 60, 50, 15, "Ready", open_button
  Text 10, 90, 55, 10, "Month and Year"
  Text 70, 90, 50, 10, "Gross Income"
  Text 130, 80, 90, 10, "Exclude from SNAP Budget"
  Text 130, 90, 30, 10, "Amount"
  Text 190, 90, 30, 10, "Reason"
  EditBox 10, 105, 40, 15, month_year
  EditBox 70, 105, 50, 15, gross_income
  EditBox 130, 105, 50, 15, exclude_from_SNAP
  EditBox 190, 105, 185, 15, exclude_reason
  Text 10, 140, 35, 10, "Tax Year"
  Text 60, 130, 35, 20, "Months in Business"
  Text 110, 140, 30, 10, "Income"
  Text 155, 140, 35, 10, "Expenses"
  EditBox 10, 155, 40, 15, tax_year
  DropListBox 60, 155, 40, 45, "12"+chr(9)+"11"+chr(9)+"10"+chr(9)+"9"+chr(9)+"8"+chr(9)+"7"+chr(9)+"6"+chr(9)+"5"+chr(9)+"4"+chr(9)+"3"+chr(9)+"2"+chr(9)+"1", months_covered
  EditBox 110, 155, 40, 15, tax_income
  EditBox 155, 155, 40, 15, tax_expenses
  ButtonGroup ButtonPressed
    PushButton 320, 155, 15, 15, "+", plus_button
    PushButton 340, 155, 15, 15, "-", minus_button
    OkButton 375, 155, 50, 15
    CancelButton 430, 155, 50, 15
EndDialog



'NAVIGATE to RBIC for each HH MEMBER and ask if Income Information was received for this RBIC
