'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - ACCT PANEL UPDATER.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at 0 because each iteration of the loop which counts the dail messages adds 1 to the counter.
STATS_manualtime = 60          'manual run time in seconds
STATS_denomination = "I"       'I is for each ITEM
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog ACCT_UPDATERone, 0, 0, 146, 150, "ACCT UPDATER"
  EditBox 64, 4, 60, 16, MAXIS_case_number
  EditBox 64, 24, 20, 16, MAXIS_footer_month
  EditBox 104, 24, 20, 16, MAXIS_footer_year
  EditBox 84, 44, 26, 16, MEMB_num
  EditBox 70, 64, 66, 16, worker_signature
  DropListBox 20, 114, 110, 16, "Select"+chr(9)+"Update"+chr(9)+"Add", add_or_update
  ButtonGroup ButtonPressed
    OkButton 4, 134, 50, 16
    CancelButton 90, 134, 50, 16
  Text 4, 70, 60, 10, "Worker Signature:"
  Text 10, 24, 50, 16, "Footer Month/Year:"
  Text 4, 50, 80, 10, "HHLD member number:"
  Text 10, 10, 50, 10, "Case Number:"
  Text 94, 30, 6, 10, "/"
  Text 4, 84, 130, 26, "Select if you are going to be updating an existing ACCT panel or adding a new ACCT panel."
EndDialog


'***This dialog will not work in dialog editor due to the ACCT type being too long for the editor to handle.
BeginDialog ACCT_UPDATERtwo, 0, 0, 246, 270, "ACCT UDATER Dialog"
  DropListBox 90, 25, 115, 15, "Select"+chr(9)+"SV - Savings"+chr(9)+"CK - Checking"+chr(9)+"CE - Certificate of Deposit"+chr(9)+"MM - Money Market"+chr(9)+"DC - Debit Card"+chr(9)+"KO - Keogh Acct"+chr(9)+"FT - Federal Thrift Savings Plan"+chr(9)+"SL - State * Local Govt Retirement and Certain Tax_Exemp Entities"+chr(9)+"RA - Employee Retirement Annuities"+chr(9)+"NP - Non-Profit Employer Ret Plans"+chr(9)+"IR - Individual Retirement Acct"+chr(9)+"RH - Roth IRA"+chr(9)+"FR - Retirement Plans for Certain Govt & Non-Govt Employers"+chr(9)+"CT - Corp Retirement Trust Prior to 6/25/1959"+chr(9)+"RT - Other Retirement Fund" +chr(9)+"QT - Qualified Tuition (529)"+chr(9)+"CA - Coverdell SV (530)"+chr(9)+"OE - Other Educational"+chr(9)+"OT - Other Account Type",ACCT_type
  EditBox 90, 45, 90, 15, ACCT_number
  EditBox 90, 65, 90, 15, ACCT_location
  EditBox 70, 85, 55, 15, Balance
  DropListBox 160, 85, 55, 15, "Select"+chr(9)+"1 - Bank Stmnt"+chr(9)+"2 - Agency Verif Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Doc"+chr(9)+"6 - Personal Stmnt"+chr(9)+"N - No Verif Provided", VERIF_type
  EditBox 65, 110, 55, 15, as_of_date
  EditBox 165, 110, 55, 15, date_recd
  EditBox 100, 130, 30, 15, penalty
  DropListBox 175, 130, 25, 15, " "+chr(9)+"Y"+chr(9)+"N", withdrawal_penalty_yes_no
  DropListBox 145, 150, 60, 15, " "+chr(9)+"1 - Bank Stmnt"+chr(9)+"2 - Agency Verif Form"+chr(9)+"3 - Coltrl Contact"+chr(9)+"5 - Other Doc"+chr(9)+"6 - Personal Stmnt"+chr(9)+"N - No Verif Provided", penalty_verif
  DropListBox 25, 185, 25, 15, " "+chr(9)+"Y"+chr(9)+"N", CASH_counts
  DropListBox 70, 185, 25, 15, " "+chr(9)+"Y"+chr(9)+"N", FS_counts
  DropListBox 115, 185, 25, 15, " "+chr(9)+"Y"+chr(9)+"N", HC_counts
  DropListBox 165, 185, 25, 15, " "+chr(9)+"Y"+chr(9)+"N", GRH_counts
  DropListBox 215, 185, 25, 15, " "+chr(9)+"Y"+chr(9)+"N", IVE_counts
  DropListBox 95, 205, 30, 10, " "+chr(9)+"Yes"+chr(9)+"No", joint_owner
  EditBox 180, 205, 10, 15, ratio_one
  EditBox 200, 205, 10, 15, ratio_two
  EditBox 125, 225, 20, 15, interest_month
  EditBox 160, 225, 20, 15, interest_year
  ButtonGroup ButtonPressed
    OkButton 65, 250, 50, 15
  Text 35, 50, 50, 10, "ACCT Number:"
  Text 30, 175, 175, 10, "Does the ACCT count towards the following programs:"
  Text 5, 190, 20, 10, "Cash:"
  Text 55, 190, 15, 10, "FS:"
  Text 100, 190, 15, 10, "HC:"
  Text 145, 190, 20, 10, "GRH:"
  Text 195, 190, 15, 10, "IV-E:"
  Text 135, 90, 20, 10, "Verif:"
  ButtonGroup ButtonPressed
    CancelButton 120, 250, 50, 15
  Text 35, 155, 100, 10, "Withdrawl Penalty Verification:"
  Text 25, 110, 40, 20, "As of Date: xx/xx/xx"
  Text 30, 210, 60, 10, "Joint Owner (y/n):"
  Text 35, 70, 50, 10, "ACCT Location:"
  Text 135, 210, 45, 10, "Share Ratio:"
  Text 35, 135, 65, 10, "Withdrawl Penalty:"
  Text 195, 210, 5, 10, "/"
  Text 35, 30, 40, 10, "ACCT type:"
  Text 30, 230, 95, 10, "Next Interest Date: MM/YY"
  Text 30, 90, 30, 10, "Balance:"
  Text 150, 230, 5, 10, "/"
  Text 145, 135, 25, 10, "Yes/No"
  Text 45, 10, 145, 10, "Please enter all pertinent ACCT information."
  Text 125, 110, 40, 20, "Date Rec'd: xx/xx/xx"
EndDialog

BeginDialog DoneUpdating, 0, 0, 114, 74, "DoneUpdating"
  ButtonGroup ButtonPressed
    OkButton 4, 58, 40, 12
  Text 8, 10, 98, 20, "Are there more ACCT panels that you need to update/add?"
  DropListBox 30, 36, 58, 12, "Select"+chr(9)+"Yes"+chr(9)+"No", more_updates
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""

CALL check_for_MAXIS(True)

Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Default member is member 01
MEMB_num = "01"

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = False

'sets excel_row at 1
excel_row = 1

DO

	DO
		CALL check_for_MAXIS(False)

		'assigns a value to more_updates
		more_updates = "Select"

		'defining variables to their default value to account for potential looping
		add_or_update = "Select"
		VERIF_type = "Select"
		ACCT_type = "Select"
		ACCT_number = ""
		ACCT_location = ""
		Balance = ""
		as_of_date = ""
		date_recd = ""
		penalty = ""
		withdrawal_penalty_yes_no = " "
		CASH_counts = " "
		FS_counts = " "
		HC_counts = " "
		GRH_counts = " "
		IVE_counts = " "
		joint_owner = "Select"
		ratio_one = ""
		ratio_two = ""
		interest_month = ""
		interest_year = ""
		panel_number = ""

		'starts the ACCT Panel Updater dialog
		err_msg = ""
		Dialog ACCT_UPDATERone
		'asks if you want to cancel and if "yes" is selected sends StopScript and closes excel
		If ButtonPressed = 0 then end_excel_and_script
		If MAXIS_case_number = "" then err_msg = err_msg & vbCr & "* Please enter a case number."
		If worker_signature = "" then err_msg = err_msg & vbCr & "* Please sign your case note."
		If MEMB_num = "" then err_msg = err_msg & vbCr & "* Please enter the member number to be updated."
		If add_or_update = "Select" then err_msg = err_msg & vbCr & "* Please select update or add an ACCT panel."
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""

	If add_or_update = "Add" THEN

		CALL check_for_MAXIS(False)

		'navigates to get program data to determine which count values are required
		call navigate_to_MAXIS_screen("stat", "prog")
		EMReadScreen CASH1_status, 4, 6, 74
		EMReadScreen CASH2_status, 4, 7, 74
		EMReadScreen GRH_status, 4, 9, 74
		EMReadScreen SNAP_status, 4, 10, 74
		EMReadScreen IV-E_status, 4, 11, 74
		EMReadScreen HC_status, 4, 12, 74

		'navigates to STAT/ACCT
		call navigate_to_MAXIS_screen("stat", "acct")


		'Heads into the case/curr screen, checks to make sure the case number is correct before proceeding. If it can't get beyond the SELF menu the script will stop.
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then end_excel_and_script

		'Navigates to the ACCT panel for the right person
		If MEMB_num <> "01" then
			EMWriteScreen MEMB_num, 20, 76
			EMWriteScreen "01", 20, 79
			transmit
		End if

		DO
			err_msg = ""
			'starts the ACCT Panel Updater dialog
			Dialog ACCT_UPDATERtwo
			'asks if you want to cancel and if "yes" is selected sends StopScript and closes excel
			IF ButtonPressed = 0 THEN end_excel_and_script
			'checks that an ACCT type has been selected
			IF ACCT_type = "Select" THEN err_msg = err_msg & vbCr & "You must select an ACCT type."
			'checks if the ACCT number has been entered.
			IF ACCT_number = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT number."
			'checks if the ACCT location has been entered.
			IF ACCT_location = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT location."
			'checks if the balance has been entered.
			IF Balance = "" THEN err_msg = err_msg & vbCr & "You must enter the balance."
			'checks that an verification type has been selected
			IF VERIF_type = "Select" THEN err_msg = err_msg & vbCr & "You must select an ACCT verification type."
			'checks if the as of date has been entered.
			IF as_of_date = "" THEN err_msg = err_msg & vbCr & "You must enter the as of date."
			'checks if the date rec'd has been entered.
			IF date_recd = "" THEN err_msg = err_msg & vbCr & "You must enter the date rec'd."
			'checks if there is a withdrawl penalty has been entered. **check is commented out currently.
			'IF withdrawal_penalty_yes_no = " " THEN err_msg = err_msg & vbCr & "You must enter if there is a withdrawal penalty."
			'checks if the CASH counts has been entered.
			IF (CASH_counts = " " AND CASH1_status = "ACTV") OR (CASH_counts = " " AND CASH1_status = "PEND") OR (CASH_counts = " " AND CASH2_status = "ACTV") OR (CASH_counts = " " AND CASH2_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for Cash."
			'checks if the FS counts has been entered.
			IF (FS_counts = " " AND SNAP_status = "ACTV") OR (FS_counts = " " AND SNAP_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for SNAP."
			'checks if the HC counts has been entered.
			IF (HC_counts = " " AND HC_status = "ACTV") OR (HC_counts = " " AND HC_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for Health Care."
			'checks if the GRH counts has been entered.
			IF (GRH_counts = " " AND GRH_status = "ACTV") OR (GRH_counts = " " AND GRH_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for GRH."
			'checks if the IVE counts has been entered.
			IF (IVE_counts = " " AND IV-E_status = "ACTV") OR (IVE_counts = " " AND IV-E_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for IVE."
			'checks if joint owner has been entered.
			IF joint_owner = " " THEN err_msg = err_msg & vbCr & "You must enter if the ACCT is a joint acct."
			'checks if ratio one has been entered.
			IF joint_owner = "Yes" and ratio_one = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT ratio."
			'checks if ratio two has been entered.
			IF joint_owner = "Yes" and ratio_two = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT ratio."
			'popups the list of errors that need to be fixed.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""

		EMWriteScreen "nn", 20, 79
		transmit

		EMReadScreen max_panel_check, 4, 24, 2 'This is just in case someone tries to create a 97th panel, 96 is currently the max for ACCT.
		IF max_panel_check = "ONLY" THEN script_end_procedure("There are the max number of panels for this Household member. Please review or delete and try again.")


		'Enters the ACCT type.
		EMWriteScreen ACCT_type, 06, 44
		'Enters the ACCT number.
		EMWriteScreen ACCT_number, 07, 44
		'Enters the ACCT location.
		EMWriteScreen ACCT_location, 08, 44
		'Enters the balance of the ACCT.
		EMWriteScreen Balance, 10, 46
		'Enters the verification type.
		EMWriteScreen VERIF_type, 10, 64
		'Enters the date of the verification.
		call create_MAXIS_friendly_date(as_of_date, 0, 11, 44)
		'Enters if there is a withdrawl penalty.
		EMWriteScreen penalty, 12, 46
		'Enters, yes or no, if there is a withdrawl penalty.
		EMWriteScreen withdrawal_penalty_yes_no, 12, 64
		'Enters if there is a verification of the withdrawl penalty.
		EMWriteScreen penalty_verif, 12, 72
		'Enters if the ACCT counts towards cash programs.
		EMWriteScreen CASH_counts, 14, 50
		'Enters if the ACCT counts towards the SNAP program.
		EMWriteScreen FS_counts, 14, 57
		'Enters if the ACCT counts towards the HC program.
		EMWriteScreen HC_counts, 14, 64
		'Enters if the ACCT counts towards the GRH program.
		EMWriteScreen GRH_counts, 14, 72
		'Enters if the ACCT counts towards the IV-E program.
		EMWriteScreen IVE_counts, 14, 80
		'Enters if the ACCT has a joint owner, yes or no.
		EMWriteScreen joint_owner, 15, 44
		'Enters the joint owner ratio (first half of the ratio).
		EMWriteScreen ratio_one, 15, 76
		'Enters the joint owner ratio (second half of the ratio).
		EMWriteScreen ratio_two, 15, 80
		'Enters the next anticipated date that interest accrues (month).
		EMWriteScreen interest_month, 17, 57
		'Enters the next anticipated date that interest accrues (year).
		EMWriteScreen interest_year, 17, 60
		transmit
		EMReadScreen panel_number, 8, 2, 72
			'trimming the panel number
			panel_number = trim(panel_number)
		'adds one instance to the stats counter
		STATS_counter = STATS_counter + 1

	else

	call check_for_MAXIS(false)

		'navigates to get program data to determine which count values are required
		call navigate_to_MAXIS_screen("stat", "prog")
		EMReadScreen CASH1_status, 4, 6, 74
		EMReadScreen CASH2_status, 4, 7, 74
		EMReadScreen GRH_status, 4, 9, 74
		EMReadScreen SNAP_status, 4, 10, 74
		EMReadScreen IV-E_status, 4, 11, 74
		EMReadScreen HC_status, 4, 12, 74

		'navigates to STAT/ACCT
		call navigate_to_MAXIS_screen("stat", "acct")


		'Heads into the case/curr screen, checks to make sure the case number is correct before proceeding. If it can't get beyond the SELF menu the script will stop.
		EMReadScreen SELF_check, 4, 2, 50
		If SELF_check = "SELF" then end_excel_and_script

		'Navigates to the ACCT panel for the right person
		If MEMB_num <> "01" then
			EMWriteScreen MEMB_num, 20, 76
			EMWriteScreen "01", 20, 79
			transmit
		End if


		'Checks to make sure there are ACCT panels for this member. If none exist the script will close
		EMReadScreen total_amt_of_panels, 2, 2, 78
		If total_amt_of_panels = " 0" then script_end_procedure("No ACCT panels exist for this client. Please restart the script and select ADD an ACCT panel or select the correct HHLD member number to update. The script will now stop.") & end_excel_and_script


		'If there is more than one panel, this part will grab ACCT info off of them and present it to the worker to decide which one to use.
		If total_amt_of_panels <> " 0" then
			Do
				EMReadScreen current_panel_number, 2, 2, 73
				EMReadScreen account_location, 30, 8, 44
				EMReadScreen account_type, 2, 6, 44
				account_check = MsgBox("Is this the ACCT you want to update? Account location : " & trim(replace(account_location, "_", "")), 3)
				If account_check = 2 then end_excel_and_script
				If account_check = 6 then
					account_found = True
					exit do
				END IF
				If account_check = 7 and current_panel_number = total_amt_of_panels then
					pick_a_different_household_member = MsgBox("You have run through all the possible ACCT panels for this person. If you need to select a different household member, press CANCEL and restart the script using the correct MEMB number. Pressing OK will continue the script using info from the ACCT panel you are currently on", vbOKCancel)
					IF pick_a_different_household_member = vbCancel THEN end_excel_and_script
					IF pick_a_different_household_member = vkOK THEN EXIT DO
				End if
				transmit
			Loop until current_panel_number = total_amt_of_panels

			'pulling ACCT info from MAXIS
	         	EMReadScreen ACCT_type, 2, 6, 44
	         	EMReadScreen ACCT_number, 20, 7, 44
	        	EMReadScreen ACCT_location, 20, 8, 44
	          	EMReadScreen Balance, 8, 10, 46
    	      	EMReadScreen VERIF_type, 1, 10, 64
    	      	EMReadScreen as_of_date, 8, 11, 44
    	      	as_of_date = replace(as_of_date, " ", "/")
    	      	EMReadScreen penalty, 8, 12, 46
	          	EMReadScreen withdrawal_penalty_yes_no, 1, 12, 64
	         	EMReadScreen penalty_verif, 1, 12, 72
	          	EMReadScreen CASH_counts, 1, 14, 50
	          	EMReadScreen FS_counts, 1, 14, 57
	          	EMReadScreen HC_counts, 1, 14, 64
	          	EMReadScreen GRH_counts, 1, 14, 72
	          	EMReadScreen IVE_counts, 1, 14, 80
	          	EMReadScreen joint_owner, 1, 15, 44
	          	EMReadScreen ratio_one, 1, 15, 76
	          	EMReadScreen ratio_two, 1, 15, 80
	          	EMReadScreen interest_month, 2, 17, 57
	          	EMReadScreen interest_year, 2, 17, 60
			EMReadScreen panel_number, 8, 2, 72

			'cleaning up the ACCT variables
			ACCT_number = replace(ACCT_number, "_", "")
			ACCT_number = trim(ACCT_number)
			ACCT_location = replace(ACCT_location, "_", "")
			ACCT_location = trim(ACCT_location)
			Balance = replace(Balance, "_", "")
			Balance = trim(Balance)
			VERIF_type = replace(VERIF_type, "_", "")
			VERIF_type = trim(VERIF_type)
			'Formatting balance verif
			IF VERIF_type = "1" THEN VERIF_type = "1 - Bank Stmnt"
			IF VERIF_type = "2" THEN VERIF_type = "2 - Agency Verif Form"
			IF VERIF_type = "3" THEN VERIF_type = "3 - Coltrl Contact"
			IF VERIF_type = "5" THEN VERIF_type = "5 - Other Doc"
			IF VERIF_type = "6" THEN VERIF_type = "6 - Personal Stmnt"
			IF VERIF_type = "N" THEN VERIF_type = "N - No Verif Provided"
			penalty = replace(penalty, "_", "")
			penalty = trim(penalty)
			withdrawal_penalty_yes_no = replace(withdrawal_penalty_yes_no, "_", "")
			withdrawal_penalty_yes_no = trim(withdrawal_penalty_yes_no)
			'formatting withdrawal_penalty_yes_no
			IF withdrawal_penalty_yes_no = "Y" THEN withdrawal_penalty_yes_no = "Yes"
			IF withdrawal_penalty_yes_no = "N" THEN withdrawal_penalty_yes_no = "No"
			penalty_verif = replace(penalty_verif, "_", "")
			penalty_verif = trim(penalty_verif)
			'formatting withdrawl penalty verif code
			IF penalty_verif = "1" THEN penalty_verif = "1 - Bank Stmnt"
			IF penalty_verif = "2" THEN penalty_verif = "2 - Agency Verif Form"
			IF penalty_verif = "3" THEN penalty_verif = "3 - Coltrl Contact"
			IF penalty_verif = "5" THEN penalty_verif = "5 - Other Doc"
			IF penalty_verif = "6" THEN penalty_verif = "6 - Personal Stmnt"
			IF penalty_verif = "N" THEN penalty_verif = "N - No Verif Provided"
			CASH_counts = replace(CASH_counts, "_", "")
			CASH_counts = trim(CASH_counts)
			FS_counts = replace(FS_counts, "_", "")
			FS_counts = trim(FS_counts)
			HC_counts = replace(HC_counts, "_", "")
			HC_counts = trim(HC_counts)
			GRH_counts = replace(GRH_counts, "_", "")
			GRH_counts = trim(GRH_counts)
			IVE_counts = replace(IVE_counts, "_", "")
			IVE_counts = trim(IVE_counts)
			joint_owner = replace(joint_owner, "_", "")
			joint_owner = trim(joint_owner)
			'formatting joint_owner
			IF joint_owner = "Y" THEN joint_owner = "Yes"
			IF joint_owner = "N" THEN joint_owner = "No"
			ratio_one = replace(ratio_one, "_", "")
			ratio_one = trim(ratio_one)
			ratio_two = replace(ratio_two, "_", "")
			ratio_two = trim(ratio_two)
			interest_month = replace(interest_month, "_", "")
			interest_month = trim(interest_month)
			interest_year = replace(interest_year, "_", "")
			interest_year = trim(interest_year)
			panel_number = trim(panel_number)
			'Formatting account type
			IF ACCT_type = "SV" THEN ACCT_type = "SV - Savings"
			IF ACCT_type = "CK" THEN ACCT_type = "CK - Checking"
			IF ACCT_type = "CE" THEN ACCT_type = "CE - Certificate of Deposit"
			IF ACCT_type = "MM" THEN ACCT_type = "MM - Money Market"
			IF ACCT_type = "DC" THEN ACCT_type = "DC - Debit Card"
			IF ACCT_type = "KO" THEN ACCT_type = "KO - Keogh Acct"
			IF ACCT_type = "FT" THEN ACCT_type = "FT - Federal Thrift Savings Plan"
			IF ACCT_type = "SL" THEN ACCT_type = "SL - State * Local Govt Retirement and Certain Tax_Exemp Entities"
			IF ACCT_type = "RA" THEN ACCT_type = "RA - Employee Retirement Annuities"
			IF ACCT_type = "NP" THEN ACCT_type = "NP - Non-Profit Employer Ret Plans"
			IF ACCT_type = "IR" THEN ACCT_type = "IR - Individual Retirement Acct"
			IF ACCT_type = "RH" THEN ACCT_type = "RH - Roth IRA"
			IF ACCT_type = "FR" THEN ACCT_type = "FR - Retirement Plans for Certain Govt & Non-Govt Employers"
			IF ACCT_type = "CT" THEN ACCT_type = "CT - Corp Retirement Trust Prior to 6/25/1959"
			IF ACCT_type = "RT" THEN ACCT_type = "RT - Other Retirement Fund"
			IF ACCT_type = "QT" THEN ACCT_type = "QT - Qualified Tuition (529)"
			IF ACCT_type = "CA" THEN ACCT_type = "CA - Coverdell SV (530)"
			IF ACCT_type = "OE" THEN ACCT_type = "OE - Other Educational"
			IF ACCT_type = "OT" THEN ACCT_type = "OT - Other Account Type"
			If ACCT_type <> "__" then variable_written_to = variable_written_to & ACCT_type & ACCT_number & ACCT_location & Balance & penalty & withdrawal_penalty_yes_no & penalty_verif & CASH_counts & FS_counts & HC_counts & GRH_counts & IVE_counts & joint_owner & ratio_one & ratio_two & interest_month & interest_year & "; "
		END IF

		DO
			err_msg = ""
			'starts the ACCT Panel Updater dialog
			Dialog ACCT_UPDATERtwo
			'asks if you want to cancel and if "yes" is selected sends StopScript and closes excel
			If ButtonPressed = 0 then end_excel_and_script
			cancel_confirmation
			'checks that an ACCT type has been selected
			IF ACCT_type = "  " THEN err_msg = err_msg & vbCr & "You must select an ACCT type."
			'checks if the ACCT number has been entered.
			IF ACCT_number = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT number."
			'checks if the ACCT location has been entered.
			IF ACCT_location = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT location."
			'checks if the balance has been entered.
			IF Balance = "" THEN err_msg = err_msg & vbCr & "You must enter the balance."
			'checks that an verification type has been selected
			IF VERIF_type = "Select" THEN err_msg = err_msg & vbCr & "You must select an ACCT verification type."
			'checks if the as of date has been entered.
			IF as_of_date = "" THEN err_msg = err_msg & vbCr & "You must enter the as of date."
			'checks if the date rec'd has been entered.
			IF date_recd = "" THEN err_msg = err_msg & vbCr & "You must enter the date rec'd."
			'checks if there is a withdrawl penalty has been entered. **check is commented out currently.
			'IF withdrawal_penalty_yes_no = " " THEN err_msg = err_msg & vbCr & "You must enter if there is a withdrawal penalty."
			'checks if the CASH counts has been entered.
			IF (CASH_counts = " " AND CASH1_status = "ACTV") OR (CASH_counts = " " AND CASH1_status = "PEND") OR (CASH_counts = " " AND CASH2_status = "ACTV") OR (CASH_counts = " " AND CASH2_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for Cash."
			'checks if the FS counts has been entered.
			IF (FS_counts = " " AND SNAP_status = "ACTV") OR (FS_counts = " " AND SNAP_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for SNAP."
			'checks if the HC counts has been entered.
			IF (HC_counts = " " AND HC_status = "ACTV") OR (HC_counts = " " AND HC_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for Health Care."
			'checks if the GRH counts has been entered.
			IF (GRH_counts = " " AND GRH_status = "ACTV") OR (GRH_counts = " " AND GRH_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for GRH."
			'checks if the IVE counts has been entered.
			IF (IVE_counts = " " AND IV-E_status = "ACTV") OR (IVE_counts = " " AND IV-E_status = "PEND") THEN err_msg = err_msg & vbCr & "You must enter if the ACCT counts for IVE."
			'checks if joint owner has been entered.
			IF joint_owner = " " THEN err_msg = err_msg & vbCr & "You must enter if the ACCT is a joint acct."
			'checks if ratio one has been entered.
			IF joint_owner = "Yes" and ratio_one = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT ratio."
			'checks if ratio two has been entered.
			IF joint_owner = "Yes" and ratio_two = "" THEN err_msg = err_msg & vbCr & "You must enter the ACCT ratio."
			'popups the list of errors that need to be fixed.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""

	PF9
	'checking to see if we got into edit mode.
	EMReadScreen edit_mode_check, 1, 20, 8
	If edit_mode_check = "D" then script_end_procedure("Unable to create a new JOBS panel. Check which member number you provided. Otherwise you may be in inquiry mode. If so switch to production and try again. Or try closing BlueZone.")

	'Assigns a value to Balance_cleanup to clear out the balance field prior to entering info from the script.
	Balance_cleanup = "        "

		'Enters the ACCT type.
		EMWriteScreen ACCT_type, 06, 44
		'Enters the ACCT number.
		EMWriteScreen ACCT_number, 07, 44
		'Enters the ACCT location.
		EMWriteScreen ACCT_location, 08, 44
		'Clears the old balance of the ACCT.
		EMWriteScreen Balance_cleanup, 10, 46
		'Enters the new balance of the ACCT.
		EMWriteScreen Balance, 10, 46
		'Enters the verification type.
		EMWriteScreen VERIF_type, 10, 64
		'Enters the date of the verification.
		call create_MAXIS_friendly_date(as_of_date, 0, 11, 44)
		'Enters if there is a withdrawl penalty.
		EMWriteScreen penalty, 12, 46
		'Enters, yes or no, if there is a withdrawl penalty.
		EMWriteScreen withdrawal_penalty_yes_no, 12, 64
		'Enters if there is a verification of the withdrawl penalty.
		EMWriteScreen penalty_verif, 12, 72
		'Enters if the ACCT counts towards cash programs.
		EMWriteScreen CASH_counts, 14, 50
		'Enters if the ACCT counts towards the SNAP program.
		EMWriteScreen FS_counts, 14, 57
		'Enters if the ACCT counts towards the HC program.
		EMWriteScreen HC_counts, 14, 64
		'Enters if the ACCT counts towards the GRH program.
		EMWriteScreen GRH_counts, 14, 72
		'Enters if the ACCT counts towards the IV-E program.
		EMWriteScreen IVE_counts, 14, 80
		'Enters if the ACCT has a joint owner, yes or no.
		EMWriteScreen joint_owner, 15, 44
		'Enters the joint owner ratio (first half of the ratio).
		EMWriteScreen ratio_one, 15, 76
		'Enters the joint owner ratio (second half of the ratio).
		EMWriteScreen ratio_two, 15, 80
		'Enters the next anticipated date that interest accrues (month).
		EMWriteScreen interest_month, 17, 57
		'Enters the next anticipated date that interest accrues (year).
		EMWriteScreen interest_year, 17, 60
		transmit
	End if

	DO
		err_msg = ""
		'starts the DoneUpdating dialog
		Dialog DoneUpdating
		If more_updates = "Select" then err_msg = err_msg & vbCr & "* Please select if there are more ACCT panels to update/add.."
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""

	IF ACCT_type = "SV" THEN ACCT_type = "SV - Savings"
	IF ACCT_type = "CK" THEN ACCT_type = "CK - Checking"
	IF ACCT_type = "CE" THEN ACCT_type = "CE - Certificate of Deposit"
	IF ACCT_type = "MM" THEN ACCT_type = "MM - Money Market"
	IF ACCT_type = "DC" THEN ACCT_type = "DC - Debit Card"
	IF ACCT_type = "KO" THEN ACCT_type = "KO - Keogh Account"
	IF ACCT_type = "FT" THEN ACCT_type = "FT - Federal Thrift Savings Plan"
	IF ACCT_type = "SL" THEN ACCT_type = "SL - State & Local Govt Ret"
	IF ACCT_type = "RA" THEN ACCT_type = "RA - Employee Ret Annuities"
	IF ACCT_type = "NP" THEN ACCT_type = "NP - Non-Profit Employer Ret Plans"
	IF ACCT_type = "IR" THEN ACCT_type = "IR - Indiv Ret Acct"
	IF ACCT_type = "RH" THEN ACCT_type = "RH - Roth IRA"
	IF ACCT_type = "FR" THEN ACCT_type = "FR - Ret Plans for Certain Govt & Non-Govt"
	IF ACCT_type = "CT" THEN ACCT_type = "CT - Corp Ret Trust Prior to 6/25/1959"
	IF ACCT_type = "RT" THEN ACCT_type = "RT - Other Ret Fund"
	IF ACCT_type = "QT" THEN ACCT_type = "QT - Qualified Tuition (529)"
	IF ACCT_type = "CA" THEN ACCT_type = "CA - Coverdell SV (530)"
	IF ACCT_type = "OE" THEN ACCT_type = "OE - Other Educational"
	IF ACCT_type = "OT" THEN ACCT_type = "OT - Other Account Type"

	'redefining variables for case note clarity
	IF add_or_update = "Add" THEN add_or_update = "added for"
	IF add_or_update = "Update" THEN add_or_update = "updated for"
	IF VERIF_type = "1" THEN VERIF_type = "1 - Bank Statement"
	IF VERIF_type = "2" THEN VERIF_type = "2 - Agency Verified"
	IF VERIF_type = "3" THEN VERIF_type = "3 - Colateral Contact"
	IF VERIF_type = "5" THEN VERIF_type = "5 - Other Doc"
	IF VERIF_type = "6" THEN VERIF_type = "6 - Personal Statement (DHS6054)"
	ACCT_location = UCase(ACCT_location)


	'the following dumps all case note info into an excel spreadsheet
	ObjExcel.Cells(excel_row, 1).Value = "* ACCT panel " & panel_number & " " & add_or_update & " Memb " & MEMB_num & ", " & ACCT_type & " at " & ACCT_location & "."
	'updates the excel row number to the next row
	excel_row = excel_row + 1
	ObjExcel.Cells(excel_row, 1).Value = "* As of " & as_of_date & ". Verif rec'd on " & date_recd & "." & " Verif type: " & VERIF_type & "."
	'updates the excel row number to the next row
	excel_row = excel_row + 1
	ObjExcel.Cells(excel_row, 1).Value = "---"

	'adds one instance to the stats counter
	STATS_counter = STATS_counter + 1

	IF more_updates = "Yes" THEN excel_row = excel_row + 1

LOOP UNTIL more_updates = "No"


'starts a blank case note
call start_a_blank_case_note

'this enters the actual case note info
IF excel_row = 3 THEN call write_variable_in_CASE_NOTE("***ACCT panel " & add_or_update & " for " & "Member " & MEMB_num & "***")
IF excel_row > 3 THEN call write_variable_in_CASE_NOTE("***Multiple ACCT panels added/updated.***")

	'assigning a value to row_num
	row_num = 1
	'looping to pull all info from excel into the case note
	DO
    		note_excel = ObjExcel.Cells(row_num, 1).Value
		call write_variable_in_CASE_NOTE(note_excel)
		row_num = row_num + 1
		excel_row = excel_row - 1
	Loop until excel_row = 0

call write_variable_in_CASE_NOTE(worker_signature)

'Manually closing workbooks so that the stats script can finish up
objExcel.Workbooks.Close
objExcel.quit

script_end_procedure("")
