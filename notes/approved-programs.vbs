'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPROVED PROGRAMS.vbs"
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
            script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
            script_repository = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/"
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
        script_repository = "C:\MAXIS-Scripts\"
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
call changelog_update("05/01/2023", "* * * THIS SCRIPT IS BEING RETIRED ON 05/08/2023 * * *##~####~##Be sure to try using NOTES - Eligibility Summary before this retirement date for CASE/NOTEs on denials. This is the time to become accustomeed to the functionality of NOTES - Eligibility Summary.##~##", "Casey Love, Hennepin County")
call changelog_update("01/13/2021", "Added temporary checkbox to case note 15% food benefit increase. Removed SNAP Banked Months case noting options.", "Ilse Ferris, Hennepin County")
call changelog_update("03/12/2020", "Removed coding specific to Banked Months that was preventing the script from continuing.", "Casey Love, Hennepin County")
call changelog_update("01/25/2019", "Removed enhanced Banked Months case noting as this is tracked by QI staff, Banked Months indicator is still within the approval note for reflecing a SNAP Banked Months case.", "Casey Love, Hennepin County")
call changelog_update("05/19/2018", "Added 'Verifs Needed' as a mandatory field for cases identified as expedited SNAP.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

get_county_code 'Checks for county info from global variables, or asks if it is not already defined.
EMConnect "" 'connecting to MAXIS
call MAXIS_case_number_finder(MAXIS_case_number) 'Finds the case number

EMReadScreen on_SELF, 4, 2, 50  'Finds the benefit month
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", bene_month, 2)
	IF bene_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & bene_month & " ", bene_year, 2)
ELSE
	CALL find_variable("Month: ", bene_month, 2)
	IF bene_month <> "" THEN CALL find_variable("Month: " & bene_month & " ", bene_year, 2)
END IF

'Converts the variables in the dialog into the variables "bene_month" and "bene_year" to autofill the edit boxes.
start_mo = bene_month
start_yr = bene_year
autofill_check = checked
'TODO Identify a Banked Months Case
'Displays the dialog and navigates to case note
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 316, 235, "Benefits Approved"
  CheckBox 80, 5, 30, 10, "SNAP", snap_approved_check
  CheckBox 115, 5, 30, 10, "Cash", cash_approved_check
  CheckBox 150, 5, 50, 10, "Health Care", hc_approved_check
  CheckBox 210, 5, 50, 10, "Emergency", emer_approved_check
  EditBox 60, 20, 55, 15, MAXIS_case_number
  ComboBox 180, 20, 80, 15, "Initial"+chr(9)+"Renewal"+chr(9)+"Recertification"+chr(9)+"Change"+chr(9)+"Reinstate", type_of_approval
  EditBox 115, 45, 195, 15, benefit_breakdown
  CheckBox 5, 65, 255, 10, "Check here to have the script autofill the approval amounts.", autofill_check
  EditBox 175, 80, 15, 15, start_mo
  EditBox 190, 80, 15, 15, start_yr
  EditBox 55, 100, 255, 15, other_notes
  EditBox 75, 120, 235, 15, programs_pending
  EditBox 55, 140, 255, 15, docs_needed
  CheckBox 10, 165, 250, 10, "Check here if SNAP was approved expedited with postponed verifications.", postponed_verif_check
  CheckBox 10, 180, 125, 10, "Check here if the case was FIATed.", FIAT_checkbox
  CheckBox 10, 195, 185, 10, "Case approved due to 15% increase in food benefits.", covid_increase_checkbox
  EditBox 75, 210, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 210, 50, 15
    CancelButton 260, 210, 50, 15
  Text 10, 85, 160, 10, "Select the first month of approval (MM YY)..."
  Text 5, 105, 45, 10, "Other Notes:"
  Text 5, 125, 70, 10, "Pending Program(s):"
  Text 5, 145, 50, 10, "Verifs Needed:"
  Text 10, 215, 60, 10, "Worker Signature: "
  Text 120, 25, 60, 10, "Type of Approval:"
  Text 5, 5, 70, 10, "Approved Programs:"
  Text 5, 25, 50, 10, "Case Number:"
  Text 5, 40, 110, 20, "Benefit Breakdown (Issuance/Spenddown/Premium):"
EndDialog

elig_summ_option_given = False
Do
	Do
		'Adding err_msg handling
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

        Call validate_MAXIS_case_number(err_msg, "*")
        If err_msg = "" and cash_approved_check = checked Then Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
        offer_test_script = True
        If dwp_status = "ACTIVE" Then offer_test_script = False
        If dwp_status = "APP CLOSE" Then offer_test_script = False
        If dwp_status = "APP OPEN" Then offer_test_script = False

        If offer_test_script = True and elig_summ_option_given = False Then
            elig_summ_option_given = True
            run_elig_summ = MsgBox("* * * THIS SCRIPT WILL BE RETIRED ON 5/8/2023 * * *" & vbCr & "Please start using Eligibility Summary for approvals of eligibile benefits right away." & vbCr & vbCr &"Run NEW Script - NOTES - Eligibliity Summary?"& vbCr & vbCr & "It appears you are running 'NOTES - Denied Programs' on a case that may be supported by the new script 'NOTES - Eligibility Summary', it is available to use to document the eligibility results denials on SNAP, CASH, HC, and EMER." & vbCr & vbCr & "The script can redirect to run NOTES - Eligibility Summary now. Remember this new script takes some time to gather the details of the approval, but reqquires little input." & vbCr & vbCr & "NOTE: Information entered in this first dialog will NOT carry through." & vbCr & vbCr & "Would you like the script to run NOTES - Eligibility Summary for you now?", vbQuestion + vbYesNo, "Redirect to NOTES - Eligibility Summary")
            If run_elig_summ = vbYes then
                script_url = script_repository & "notes\eligibility-summary.vbs"
                ' MsgBox script_url
                Call run_from_GitHub(script_url)
            End If
        End If

		'Enforcing mandatory fields
		IF autofill_check = checked THEN
			IF snap_approved_check = unchecked AND cash_approved_check = unchecked AND emer_approved_check = unchecked THEN err_msg = err_msg & _
			 vbCr & "* You checked to have the approved amount autofilled but have not selected a program with an approval amount. Please check your selections."
		End If
        If postponed_verif_check = checked and trim(docs_needed) = "" then err_msg = err_msg & vbCr & "* Please enter the postponed verifications needed/requested."
        'If SNAP_banked_mo_check = checked AND (trim(banked_footer_month) = "" OR trim(banked_footer_year) = "") Then err_msg = err_msg & vbNewLine & "* Indicate the first month being approved with BANKED MONTHS since this approval includes a BANKED MONTH."
		IF worker_signature = "" then err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
Loop until are_we_passworded_out = false

Call date_array_generator (start_mo, start_yr, date_array)

'TODO add constants for the array
Dim BENE_AMOUNT_ARRAY()	'Array to store all the different elig amounts
Redim BENE_AMOUNT_ARRAY(reporter_type, 0)

Const progs_to_check     = 0
Const benefit_month      = 1
Const benefit_year       = 2
Const snap_amount        = 3
Const case_prorated_date = 4
Const mfip_cash          = 5
Const mfip_housing_grant = 6
Const dwp_shelter        = 7
Const dwp_personal       = 8
Const other_cash         = 9
Const reporter_type      = 10

DIM ALL_SNAP_CLIENTS_ARRAY()	'Array to check clients for ABAWD
ReDim ALL_SNAP_CLIENTS_ARRAY(banked_months_approved,0)

Const clt_ref_nbr           = 0
Const client_name           = 1
Const client_age            = 2
Const client_fset_status    = 3
Const wreg_status           = 4
Const using_banked_check    = 5
Const initial_banked_month  = 6
Const initial_banked_year   = 7
Const banked_months_approved   = 8

Dim BM_Clients_Array () 	'Array of all clients approved for BANKED MONTHS with this approval
clt_banked_mo_apprvd = 0	'g

all_elig_results = 0

'Gathers all programs with elig results from ELIG SUMM and adds them to an array
'The array is per elig program and month
'TODO - look for multiple cash programs - this doesn't work if DWP is being closed and MFIP is being opened.
For each item in date_array
	Call navigate_to_MAXIS_screen("ELIG", "SUMM")
	cur_month = datepart("m", item)
	If len(cur_month) = 1 then cur_month = "0" & cur_month
	cur_year = right(datepart("yyyy", item), 2)
	EMWriteScreen cur_month, 19, 56
	EMWriteScreen cur_year, 19, 59
	transmit
	For row = 7 to 18
		EMReadScreen versions_exist, 1, row, 40
		If versions_exist <> " " THEN
			EMReadScreen version_date, 8, row, 48
			If cdate(version_date) = date THEN
				Redim Preserve BENE_AMOUNT_ARRAY(reporter_type, all_elig_results)
				EMReadScreen prog_to_check, 4, row, 22
				'EMReadScreen snap_month, 2, 19, 56
				'EMReadScreen snap_year, 2, 19, 59
				prog_to_check = trim(prog_to_check)
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = prog_to_check
				BENE_AMOUNT_ARRAY(benefit_month, all_elig_results) = cur_month
				BENE_AMOUNT_ARRAY(benefit_year, all_elig_results) = cur_year
				all_elig_results = all_elig_results + 1
			End If
		End If
	Next
Next

infant_on_case = "Unknown"
months_of_benes = 0

'Here the script will use the program listed in the array to determine where to go to find the amounts - then add them to the array
For all_elig_results = 0 to UBound (BENE_AMOUNT_ARRAY,2)
    If postponed_verif_check = checked AND BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "Food" Then
        If xfs_package <> "" Then
            If months_of_benes >= xfs_package Then BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
        End If
    End If

	If BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "Food" AND snap_approved_check = checked Then

        banked_month_case = FALSE
        banked_month_counter = ""

        back_to_self
        MAXIS_footer_month = BENE_AMOUNT_ARRAY(benefit_month, all_elig_results)
        MAXIS_footer_year = BENE_AMOUNT_ARRAY(benefit_year, all_elig_results)
        navigate_to_MAXIS_screen "STAT", "WREG"
        stat_row = 5
        Do
            EMReadScreen memb_ref_numb, 2, stat_row, 3
            If memb_ref_numb = "  " Then Exit Do
            EMWRiteScreen memb_ref_numb, 20, 76
            transmit

            EMReadScreen fset_code, 2, 8, 50
            EMReadScreen abawd_code, 2, 13, 50
            EMReadScreen banked_code, 1, 14, 50

            If abawd_code = "13" Then banked_month_case = TRUE
            If banked_code <> "_" Then banked_month_counter = banked_code

            stat_row = stat_row + 1
        Loop until stat_row = 20

		Call navigate_to_MAXIS_screen("ELIG", "FS")
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_month, all_elig_results), 19, 54
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_year, all_elig_results), 19, 57
		EMWRiteScreen "FSSM", 19, 70
		transmit
		EMReadScreen notc_type, 8, 3, 3
		If notc_type = "APPROVED" then
			EMReadScreen snap_bene_amt, 8, 13, 73
			EMReadScreen snap_reporter, 10, 8, 31
			EMReadScreen partial_bene, 8, 9, 44
			If partial_bene = "Prorated" then
                ' If banked_month_case = TRUE and banked_month_counter <> "" Then
                '     end_message = "This is a Banked Months SNAP case." & vbNewLine & BENE_AMOUNT_ARRAY(benefit_month, all_elig_results) & "/" & BENE_AMOUNT_ARRAY(benefit_year, all_elig_results) & " is a prorated month." &_
                '     vbNewLine & "WREG has Banked Month counted to be - " & banked_month_counter & " in this footer month." &_
                '     vbNewLine & "A Banked Month should not be counted in a prorated month."
                '     script_end_procedure(end_message)
                ' End If
                EMReadScreen prorated_date, 8, 9, 58
				BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = prorated_date
                day_of_proration = DatePart("d", prorated_date)
                If day_of_proration < 15 Then
                    xfs_package = 1
                Else
                    xfs_package = 2
                End If
			End If
			BENE_AMOUNT_ARRAY(snap_amount, all_elig_results) = trim(snap_bene_amt)
			BENE_AMOUNT_ARRAY(reporter_type, all_elig_results) = snap_reporter & " Reporter"
		ELSE
			EMReadScreen approval_versions, 2, 2, 18
			If trim(approval_versions) = "1" THEN
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
				IF snap_approved_check = checked THEN MsgBox "This is not an approved version from today, SNAP amounts will not be case noted"
			Else
				approval_versions = abs(approval_versions)
				approval_to_check = approval_versions - 1
				EMWriteScreen approval_to_check, 19, 78
				transmit
				EMReadScreen approval_date, 8, 3, 14
				approval_date = cdate(approval_date)
				If approval_date = date THEN
					EMReadScreen snap_bene_amt, 8, 13, 73
					EMReadScreen snap_reporter, 10, 8, 31
					EMReadScreen partial_bene, 8, 9, 44
					If partial_bene = "Prorated" then
                        If banked_month_case = TRUE and banked_month_counter <> "" Then
                            end_message = "This is a Banked Months SNAP case." & vbNewLine & BENE_AMOUNT_ARRAY(benefit_month, all_elig_results) & "/" & BENE_AMOUNT_ARRAY(benefit_year, all_elig_results) & " is a prorated month." &_
                            vbNewLine & "WREG has Banked Month counted to be - " & banked_month_counter & " in this footer month." &_
                            vbNewLine & "A Banked Month should not be counted in a prorated month."
                            script_end_procedure(end_message)
                        End If
						EMReadScreen prorated_date, 8, 9, 58
						BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = prorated_date
                        day_of_proration = DatePart("d", prorated_date)
                        If day_of_proration < 15 Then
                            xfs_package = 1
                        Else
                            xfs_package = 2
                        End If
					End If
					BENE_AMOUNT_ARRAY(snap_amount, all_elig_results) = trim(snap_bene_amt)
					BENE_AMOUNT_ARRAY(reporter_type, all_elig_results) = trim(snap_reporter) & " Reporter"
				Else
					IF snap_approved_check = checked THEN MsgBox "Your most recent SNAP approval for the benefit month chosen is not from today. This approval amount will not be case noted"
					BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
				End If
			End If
		End If
        months_of_benes = months_of_benes + 1
	ElseIf BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "MFIP" AND cash_approved_check = checked Then
		If infant_on_case = "Unknown" Then
			Call navigate_to_MAXIS_screen ("STAT", "PNLP")
			pnlp_row = 3
			Do
				EMReadScreen panel_name, 4, pnlp_row, 5
				If panel_name = "MEMB" Then
					EMReadScreen clt_age, 2, pnlp_row, 71
					If clt_age = " 0" Then
						infant_on_case = TRUE
						Exit Do
					End If
				ElseIf panel_name = "MEMI" Then
					infant_on_case = FALSE
					Exit Do
				End IF
				pnlp_row = pnlp_row + 1
				If pnlp_row = 20 Then
					transmit
					pnlp_row = 3
				End If
			Loop Until panel_name = "REVW"
		End If
		Call navigate_to_MAXIS_screen("ELIG", "MFIP")
		'Checking that the MFIP case does not have a significant change determination page (ELIG/MFSC). We need to transmit through that page to get to ELIG/MFPR.
		row = 1
		col = 1
		EMSearch "(MFSC)", row, col
		IF row <> 0 THEN transmit
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_month, all_elig_results), 20, 56
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_year, all_elig_results), 20, 59
		EMWriteScreen "MFSM", 20, 71
		transmit
		EMReadScreen cash_approved_version, 8, 3, 3
		If cash_approved_version = "APPROVED" Then
			EMReadScreen cash_approval_date, 8, 3, 14
			If cdate(cash_approval_date) = date Then
				EMReadScreen mfip_bene_cash_amt, 8, 14, 73
				EMReadScreen mfip_bene_food_amt, 8, 15, 73
				EMReadScreen mfip_bene_housing_amt, 8, 16, 73
				EMReadScreen mfip_reporter, 10, 8, 31
				EMWriteScreen "MFB2", 20, 71
				transmit
				EMReadScreen prorate_date, 8, 5, 19
				BENE_AMOUNT_ARRAY(mfip_cash, all_elig_results) = trim(mfip_bene_cash_amt)
				BENE_AMOUNT_ARRAY(snap_amount, all_elig_results) = trim(mfip_bene_food_amt)
				BENE_AMOUNT_ARRAY(mfip_housing_grant, all_elig_results) = trim(mfip_bene_housing_amt)
				BENE_AMOUNT_ARRAY(reporter_type, all_elig_results) = trim(mfip_reporter) & " Reporter"
				If prorate_date <> "        " Then BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = prorate_date
			Else
				IF cash_approved_check = checked THEN MsgBox "This MFIP approval was not done today and the benefit amount will not be case noted"
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			End If
		Else
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of CASH in the selected benefit month. Please approve before running the script."
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			Else
				cash_approval_versions = abs(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 79
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				IF cdate(cash_approval_date) = date THEN
					EMReadScreen mfip_bene_cash_amt, 8, 14, 73
					EMReadScreen mfip_bene_food_amt, 8, 15, 73
					EMReadScreen mfip_bene_housing_amt, 8, 16, 73
					EMReadScreen mfip_reporter, 10, 8, 31
					EMWriteScreen "MFB2", 20, 71
					transmit
					EMReadScreen prorate_date, 8, 5, 19
					BENE_AMOUNT_ARRAY(mfip_cash, all_elig_results) = trim(mfip_bene_cash_amt)
					BENE_AMOUNT_ARRAY(snap_amount, all_elig_results) = trim(mfip_bene_food_amt)
					BENE_AMOUNT_ARRAY(mfip_housing_grant, all_elig_results) = trim(mfip_bene_housing_amt)
					BENE_AMOUNT_ARRAY(reporter_type, all_elig_results) = trim(mfip_reporter) & " Reporter"
					If prorate_date <> "        " Then BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = prorate_date
				Else
					IF cash_approved_check = checked THEN MsgBox "Your most recent MFIP approval is not from today and benefit amounts will not be added to case note"
					BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
				End If
			End If
		End If
	ElseIf BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "DWP" AND cash_approved_check = checked THEN
		If infant_on_case = "Unknown" Then
			Call navigate_to_MAXIS_screen ("STAT", "PNLP")
			pnlp_row = 3
			Do
				EMReadScreen panel_name, 4, pnlp_row, 5
				If panel_name = "MEMB" Then
					EMReadScreen clt_age, 2, pnlp_row, 71
					If clt_age = " 0" Then
						infant_on_case = TRUE
						Exit Do
					End If
				ElseIf panel_name = "MEMI" Then
					infant_on_case = FALSE
					Exit Do
				End IF
				pnlp_row = pnlp_row + 1
				If pnlp_row = 20 Then
					transmit
					pnlp_row = 3
				End If
			Loop Until panel_name = "REVW"
		End If
		Call navigate_to_MAXIS_screen("ELIG", "DWP")
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_month, all_elig_results), 20, 56
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_year, all_elig_results), 20, 59
		EMWriteScreen "DWSM", 20, 71
		transmit
		EMReadScreen cash_approved_version, 8, 3, 3
		If cash_approved_version = "APPROVED" Then
			EMReadScreen cash_approval_date, 8, 3, 14
			If cdate(cash_approval_date) = date Then
				EMReadScreen DWP_bene_shel_amt, 8, 13, 73
				EMReadScreen DWP_bene_pers_amt, 8, 14, 73
				EMWriteScreen "DWB2", 20, 71
				transmit
				EMReadScreen prorate_date, 8, 6, 18
				BENE_AMOUNT_ARRAY(dwp_shelter, all_elig_results) = trim(DWP_bene_shel_amt)
				BENE_AMOUNT_ARRAY(dwp_personal, all_elig_results) = trim(DWP_bene_pers_amt)
				IF prorate_date <> "__ __ __" Then BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = Replace(prorate_date, " ", "/")
			Else
				IF cash_approved_check = checked THEN MsgBox "This DWP approval was not done today and the benefit amount will not be case noted"
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			End If
		Else
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of CASH in the selected benefit month. Please approve before running the script."
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			Else
				cash_approval_versions = abs(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 79
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				If cdate(cash_approval_date) = date Then
					EMReadScreen DWP_bene_shel_amt, 8, 13, 73
					EMReadScreen DWP_bene_pers_amt, 8, 14, 73
					EMWriteScreen "DWB2", 20, 71
					transmit
					EMReadScreen prorate_date, 8, 6, 18
					'Add prorated information gathering
					BENE_AMOUNT_ARRAY(dwp_shelter, all_elig_results) = trim(DWP_bene_shel_amt)
					BENE_AMOUNT_ARRAY(dwp_personal, all_elig_results) = trim(DWP_bene_pers_amt)
					IF prorate_date <> "__ __ __" Then BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = Replace(prorate_date, " ", "/")
				Else
					IF cash_approved_check = checked THEN MsgBox "Your most recent DWP approval is not from today and benefit amounts will not be added to case note"
					BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
				End If
			End If
		End If
	ElseIf BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "GA" AND cash_approved_check = checked THEN
		'GA portion
		call navigate_to_MAXIS_screen("ELIG", "GA")
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_month, all_elig_results), 20, 54
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_year, all_elig_results), 20, 57
		EMWRiteScreen "GASM", 20, 70
		transmit
		EMReadScreen cash_approved_version, 8, 3, 3
		IF cash_approved_version = "APPROVED" THEN
			EMReadScreen cash_approval_date, 8, 3, 15
			IF cdate(cash_approval_date) = date THEN
				EMReadScreen GA_bene_cash_amt, 8, 14, 72
				EMWriteScreen "GAB2", 20, 70
				transmit
				EMReadScreen prorate_date, 5, 10, 14
				BENE_AMOUNT_ARRAY(other_cash, all_elig_results) = trim(GA_bene_cash_amt)
				IF prorate_date <> "     " Then BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = Replace(prorate_date, " ", "/") & "/" & BENE_AMOUNT_ARRAY(benefit_year,all_elig_results)
			Else
				IF cash_approved_check = checked THEN MsgBox "The most recent approval is not from today and will not be added to the case note"
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			END IF
		ELSE
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of GA in the selected benefit month. This will not be added to the case note."
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			Else
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 78
				transmit
				EMReadScreen cash_approval_date, 8, 3, 15
				IF cdate(cash_approval_date) = date THEN
					EMReadScreen GA_bene_cash_amt, 8, 14, 72
					EMWriteScreen "GAB2", 20, 70
					transmit
					EMReadScreen prorate_date, 5, 10, 14
					BENE_AMOUNT_ARRAY(other_cash, all_elig_results) = trim(GA_bene_cash_amt)
					IF prorate_date <> "     " Then BENE_AMOUNT_ARRAY(case_prorated_date, all_elig_results) = Replace(prorate_date, " ", "/") & "/" & BENE_AMOUNT_ARRAY(benefit_year,all_elig_results)
				Else
					IF cash_approved_check = checked THEN MsgBox "The most recent approval is not from today and will not be added to the case note"
					BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
				END IF
			End If
		END IF
	ELSEIF BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "MSA" AND cash_approved_check = checked THEN
		'MSA portion
		call navigate_to_MAXIS_screen("ELIG", "MSA")
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_month, all_elig_results), 20, 56
		EMWriteScreen BENE_AMOUNT_ARRAY(benefit_year, all_elig_results), 20, 59
		EMWRiteScreen "MSSM", 20, 71
		transmit
		EMReadScreen cash_approved_version, 8, 3, 3
		IF cash_approved_version = "APPROVED" THEN
			EMReadScreen cash_approval_date, 8, 3, 14
			IF cdate(cash_approval_date) = date THEN
				EMReadScreen MSA_bene_cash_amt, 8, 17, 73
				'MSA does not have proration
				BENE_AMOUNT_ARRAY(other_cash, all_elig_results) = trim(MSA_bene_cash_amt)
			Else
				IF cash_approved_check = checked THEN MsgBox "The most recent approval is not from today and will not be added to the case note"
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			END IF
		ELSE
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of MSA in the selected benefit month. This will not be added to the case note"
				BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
			Else
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 78
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				IF cdate(cash_approval_date) = date THEN
					EMReadScreen MSA_bene_cash_amt, 8, 17, 73
					'MSA does not have proration
					BENE_AMOUNT_ARRAY(other_cash, all_elig_results) = trim(MSA_bene_cash_amt)
				Else
					IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of MSA in the selected benefit month. This will not be added to the case note"
					BENE_AMOUNT_ARRAY(progs_to_check, all_elig_results) = "NONE"
				END IF
			End If
		END IF
	END IF
    Call back_to_SELF
Next

'Case notes----------------------------------------------------------------------------------------------------
IF infant_on_case = TRUE Then
	Call navigate_to_MAXIS_screen ("STAT", "EMPS")
	baby_warning = MsgBox ("This is a family cash (MFIP or DWP) case with a child under 1 year old on it." & vbNewLine & vbNewLine & "These cases are error prone particularly at intake. Please review the EMPS panel to be sure the coding matches the client request.", vbSystemModal, "Child Under 1 Year Cash Case Warning")
End If

call start_a_blank_CASE_NOTE	'Case note for the general approval
IF snap_approved_check = checked THEN
	IF postponed_verif_check = checked THEN
		approved_programs = approved_programs & "EXPEDITED SNAP/"
	ELSE
		approved_programs = approved_programs & "SNAP/"
	END IF
    If SNAP_banked_mo_check = checked Then approved_programs = "BANKED " & approved_programs
END IF

IF hc_approved_check = checked THEN approved_programs = approved_programs & "HC/"
IF cash_approved_check = checked THEN approved_programs = approved_programs & "CASH/"
IF emer_approved_check = checked THEN approved_programs = approved_programs & "EMER/"
EMSendKey "---Approved " & approved_programs & "<backspace>" & " " & type_of_approval & "---" & "<newline>"
IF postponed_verif_check = checked THEN write_variable_in_CASE_NOTE("**EXPEDITED SNAP APPROVED BUT CASE HAS POSTPONED VERIFICATIONS.**")
IF benefit_breakdown <> "" THEN call write_bullet_and_variable_in_case_note("Benefit Breakdown", benefit_breakdown)
IF autofill_check = checked THEN
	FOR snap_approvals = 0 to UBound(BENE_AMOUNT_ARRAY,2)
		IF BENE_AMOUNT_ARRAY(progs_to_check,snap_approvals) = "Food" AND snap_approved_check = checked THEN
			snap_header = ("SNAP for " & BENE_AMOUNT_ARRAY(benefit_month,snap_approvals) & "/" & BENE_AMOUNT_ARRAY(benefit_year,snap_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE (snap_header, FormatCurrency(BENE_AMOUNT_ARRAY(snap_amount,snap_approvals)) & " " & BENE_AMOUNT_ARRAY(reporter_type,snap_approvals))
			IF BENE_AMOUNT_ARRAY(case_prorated_date, snap_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", BENE_AMOUNT_ARRAY(case_prorated_date,snap_approvals))
			End If
		End If
	Next
	FOR mfip_approvals = 0 to UBound(BENE_AMOUNT_ARRAY,2)
		IF BENE_AMOUNT_ARRAY(progs_to_check,mfip_approvals) = "MFIP" AND cash_approved_check = checked THEN
			Call write_variable_in_CASE_NOTE ("MFIP for " & BENE_AMOUNT_ARRAY(benefit_month,mfip_approvals) & "/" & BENE_AMOUNT_ARRAY(benefit_year,mfip_approvals) & " " & BENE_AMOUNT_ARRAY(reporter_type,mfip_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE ("Cash Portion", FormatCurrency(BENE_AMOUNT_ARRAY(mfip_cash, mfip_approvals)))
			Call write_bullet_and_variable_in_CASE_NOTE ("Food Portion", FormatCurrency(BENE_AMOUNT_ARRAY(snap_amount, mfip_approvals)))
			Call write_bullet_and_variable_in_CASE_NOTE ("Housing Grant Amount", FormatCurrency(BENE_AMOUNT_ARRAY(mfip_housing_grant, mfip_approvals)))
			IF BENE_AMOUNT_ARRAY(case_prorated_date, mfip_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", BENE_AMOUNT_ARRAY(case_prorated_date,mfip_approvals))
			End If
		End If
	Next
	FOR dwp_approvals = 0 to UBound(BENE_AMOUNT_ARRAY,2)
		IF BENE_AMOUNT_ARRAY(progs_to_check,dwp_approvals) = "DWP" AND cash_approved_check = checked THEN
			Call write_variable_in_CASE_NOTE ("DWP for " & BENE_AMOUNT_ARRAY(benefit_month,dwp_approvals) & "/" & BENE_AMOUNT_ARRAY(benefit_year,dwp_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE ("Shelter Benefit", FormatCurrency(BENE_AMOUNT_ARRAY(dwp_shelter, dwp_approvals)))
			Call write_bullet_and_variable_in_CASE_NOTE ("Personal Needs", FormatCurrency(BENE_AMOUNT_ARRAY(dwp_personal, dwp_approvals)))
			IF BENE_AMOUNT_ARRAY(case_prorated_date, dwp_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", BENE_AMOUNT_ARRAY(case_prorated_date,dwp_approvals))
			End If
		End If
	Next
	FOR msa_approvals = 0 to UBound(BENE_AMOUNT_ARRAY, 2)
		IF BENE_AMOUNT_ARRAY(progs_to_check,msa_approvals) = "MSA" AND cash_approved_check = checked THEN
			msa_header = ("MSA for " & BENE_AMOUNT_ARRAY(benefit_month,msa_approvals) & "/" & BENE_AMOUNT_ARRAY(benefit_year, msa_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE (msa_header, FormatCurrency(BENE_AMOUNT_ARRAY(other_cash,msa_approvals)))
		End If
	Next
	FOR ga_approvals = 0 to UBound(BENE_AMOUNT_ARRAY, 2)
		IF BENE_AMOUNT_ARRAY(progs_to_check,ga_approvals) = "GA" AND cash_approved_check = checked THEN
			ga_header = ("GA for " & BENE_AMOUNT_ARRAY(benefit_month,ga_approvals) & "/" & BENE_AMOUNT_ARRAY(benefit_year,ga_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE (ga_header, FormatCurrency(BENE_AMOUNT_ARRAY(other_cash,ga_approvals)))
			IF BENE_AMOUNT_ARRAY(case_prorated_date, ga_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", BENE_AMOUNT_ARRAY(case_prorated_date,ga_approvals))
			End If
		End If
	Next
END IF
IF FIAT_checkbox = 1 THEN Call write_variable_in_CASE_NOTE ("* This case has been FIATed.")
IF other_notes <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Approval Notes", other_notes)
IF programs_pending <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Programs Pending", programs_pending)
If docs_needed <> "" then call write_bullet_and_variable_in_CASE_NOTE("Verifs needed", docs_needed)
IF SNAP_banked_mo_check = checked THEN Call write_variable_in_CASE_NOTE ("* BANKED MONTHS were approved on this case starting with " & banked_footer_month & "/" & banked_footer_year & ".")
If covid_increase_checkbox = 1 then Call write_variable_in_CASE_NOTE ("* Case approved due to 15% increase in food benefits.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("Success! Please remember to check the generated notice to make sure it is correct. If not, please add WCOMs to make notice read correctly.")
