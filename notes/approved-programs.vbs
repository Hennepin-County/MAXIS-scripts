'Rewrite by Casey Love from Ramsey County. Based on the original script by Robert Kalb and Charles Potter from Anoka County and and Ilse Ferris from Hennepin County. Veronica from DHS helped.

'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPROVED PROGRAMS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
'END OF stats block=========================================================================================================

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
call changelog_update("12/1/2016", "Initial version.", "David Courtwright, St Louis County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog benefits_approved, 0, 0, 271, 235, "Benefits Approved"
  CheckBox 80, 5, 30, 10, "SNAP", snap_approved_check
  CheckBox 115, 5, 30, 10, "Cash", cash_approved_check
  CheckBox 150, 5, 50, 10, "Health Care", hc_approved_check
  CheckBox 210, 5, 50, 10, "Emergency", emer_approved_check
  EditBox 60, 20, 55, 15, MAXIS_case_number
  ComboBox 180, 20, 80, 15, "Initial"+chr(9)+"Renewal"+chr(9)+"Recertification"+chr(9)+"Change"+chr(9)+"Reinstate", type_of_approval
  EditBox 115, 45, 150, 15, benefit_breakdown
  CheckBox 5, 65, 255, 10, "Check here to have the script autofill the approval amounts.", autofill_check
  EditBox 175, 80, 15, 15, start_mo
  EditBox 190, 80, 15, 15, start_yr
  EditBox 55, 100, 210, 15, other_notes
  EditBox 75, 120, 190, 15, programs_pending
  EditBox 55, 140, 210, 15, docs_needed
  CheckBox 10, 165, 250, 10, "Check here if SNAP was approved expedited with postponed verifications.", postponed_verif_check
  CheckBox 10, 180, 125, 10, "Check here if the case was FIATed", FIAT_checkbox
  CheckBox 10, 195, 190, 10, "Check here if SNAP BANKED MONTHS were approved.", SNAP_banked_mo_check
  EditBox 75, 210, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 210, 50, 15
    CancelButton 215, 210, 50, 15
  Text 5, 40, 110, 20, "Benefit Breakdown (Issuance/Spenddown/Premium):"
  Text 10, 85, 160, 10, "Select the first month of approval (MM YY)..."
  Text 5, 105, 45, 10, "Other Notes:"
  Text 5, 125, 70, 10, "Pending Program(s):"
  Text 5, 145, 50, 10, "Verifs Needed:"
  Text 10, 215, 60, 10, "Worker Signature: "
  Text 120, 25, 60, 10, "Type of Approval:"
  Text 5, 5, 70, 10, "Approved Programs:"
  Text 5, 25, 50, 10, "Case Number:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS
EMConnect ""
'Finds the case number
call MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------FUNCTIONS WE INVENTED THAT WILL SOON BE ADDED TO FUNCLIB
FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date
	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
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

'Displays the dialog and navigates to case note
Do
	Do
		'Adding err_msg handling
		err_msg = ""
		Dialog benefits_approved
		cancel_confirmation
			'Enforcing mandatory fields
			If MAXIS_case_number = "" then err_msg = err_msg & vbCr & "* Please enter a case number."
			IF autofill_check = checked THEN
				IF snap_approved_check = unchecked AND cash_approved_check = unchecked AND emer_approved_check = unchecked THEN err_msg = err_msg & _
				 vbCr & "* You checked to have the approved amount autofilled but have not selected a program with an approval amount. Please check your selections."
			End If
			IF worker_signature = "" then err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
Loop until are_we_passworded_out = false

Call date_array_generator (start_mo, start_yr, date_array)

Dim bene_amount_array ()	'Array to store all the different elig amounts
Redim bene_amount_array (10, 0)
	'bene_amount_array Map
	'(0, a) = Program to Check
	'(1, a) = Benefit Month
	'(2, a) = Benefit Year
	'(3, a) = SNAP amount
	'(4, a) = Prorated date
	'(5, a) = MFIP Cash
	'(6, a) = MFIP Housing Grant
	'(7, a) = DWP Shelter Amount
	'(8, a) = DWP Personal Amount
	'(9, a) = Other Cash amount
	'(10, a) = Type of Reporter

DIM All_SNAP_Clients_Array ()	'Array to check clients for ABAWD
ReDim All_SNAP_Clients_Array (8,0)
	'All_SNAP_Clients_Array MAP
	'(0, a) = Client Reference Number
	'(1, a) = Client Name
	'(2, a) = Client Age
	'(3, a) = FSET Status
	'(4, a) = WREG Status
	'(5, a) = using banked months check
	'(6, a) = Initial Banked Month
	'(7, a) = Initial Banked Year

Dim BM_Clients_Array () 	'Array of all clients approved for BANKED MONTHS with this approval
clt_banked_mo_apprvd = 0	'g

'If worker selects that banked months have been approved, script will write additional case notes, document the months and tikl
IF SNAP_banked_mo_check = checked THEN
	clients_on_case = 0 	'b

	navigate_to_MAXIS_screen "STAT", "MEMB"
	DO								'Gets name, ref number, and age for all clients and adds to an array
		ReDim Preserve All_SNAP_Clients_Array (8, clients_on_case)
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadscreen last_name, 25, 6, 30
		EMReadscreen first_name, 12, 6, 63
		EMReadscreen Mid_intial, 1, 6, 79
		EMReadScreen age, 2, 8, 76
		last_name = replace(last_name, "_", "") & " "
		first_name = replace(first_name, "_", "") & " "
		mid_initial = replace(mid_initial, "_", "")
		All_SNAP_Clients_Array (0, clients_on_case) = ref_nbr
		All_SNAP_Clients_Array (1, clients_on_case) = last_name & first_name & mid_intial
		All_SNAP_Clients_Array (2, clients_on_case) = trim(age)
		clients_on_case = clients_on_case + 1
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	wreg_check = 0 		'zero-ing out the value of wreg_check variable

	DO 		'Gets information from WREG for each client
		navigate_to_MAXIS_screen "STAT", "WREG"
		EMWriteScreen All_SNAP_Clients_Array(0, wreg_check), 20, 76
		transmit
		EMReadScreen FSET_status, 2, 8, 50
		EMReadScreen ABAWD_status, 2, 13, 50
		All_SNAP_Clients_Array (3, wreg_check) = FSET_status
		All_SNAP_Clients_Array (4, wreg_check) = ABAWD_status
		IF FSET_status = "30" then
			IF ABAWD_status = 10 OR ABAWD_status = 11 then
				All_SNAP_Clients_Array (5, wreg_check) = 1
				All_SNAP_Clients_Array (6, wreg_check) = start_mo
				All_SNAP_Clients_Array (7, wreg_check) = start_yr
				All_SNAP_Clients_Array (8, wreg_check) = 3
			End If
		End If
		wreg_check = wreg_check + 1
	Loop until wreg_check = clients_on_case

	'Dialog is defined here because the Array needs to happen first for it to be dynamic
	BeginDialog Banked_Months_Dialog, 0, 0, 330, ((clients_on_case * 45) + 40), "Determining Clients Using Banked Months"
	  Text 65, 5, 145, 10, "Household Members Using Banked Months"
	  For client_dialog = 0 to (clients_on_case - 1)
		CheckBox 5, (20 + (client_dialog * 45)), 85, 10, "Using Banked Months", All_SNAP_Clients_Array (5, client_dialog)
		Text 100, (20 + (client_dialog * 45)), 65, 10, "First Banked Month"
		EditBox 165, (15 + (client_dialog * 45)), 15, 15, All_SNAP_Clients_Array (6, client_dialog)
		EditBox 180, (15 + (client_dialog * 45)), 15, 15, All_SNAP_Clients_Array (7, client_dialog)
		 Text 205, (20 + (client_dialog * 45)), 100, 10, "Number of Banked Months App"
		EditBox 310, (15 + (client_dialog * 45)), 15, 15, All_SNAP_Clients_Array (8, client_dialog)
		Text 20, (35 + (client_dialog * 45)), 30, 10, "Memb " & All_SNAP_Clients_Array (0, client_dialog)
		Text 60, (35 + (client_dialog * 45)), 210, 10, All_SNAP_Clients_Array (1, client_dialog)
		Text 60, (50 + (client_dialog * 45)), 35, 10, "FSET: " & All_SNAP_Clients_Array (3, client_dialog)
		Text 100, (50 + (client_dialog * 45)), 45, 10, "ABAWD: " & All_SNAP_Clients_Array (4, client_dialog)
		Text 265, (35 + (client_dialog * 45)), 35, 10, "Age: " & All_SNAP_Clients_Array (2, client_dialog)
	  Next
	  ButtonGroup ButtonPressed
		OkButton 220, (65 + ((clients_on_case - 1) * 45)), 50, 15
		CancelButton 275, (65 + ((clients_on_case - 1) * 45)), 50, 15
	EndDialog

	Do
		err_msg = ""
		Dialog Banked_Months_Dialog
		cancel_confirmation
		clients_with_banked_mo = 0
		FOR clt_err_check = 0 to (clients_on_case - 1)
			IF All_SNAP_Clients_Array (5, clt_err_check) = checked THEN
				clients_with_banked_mo = clients_with_banked_mo + 1
				IF All_SNAP_Clients_Array (6, clt_err_check)  = "" AND All_SNAP_Clients_Array(7, clt_err_check) = "" THEN _
				err_msg = err_msg & vbCr & "You must indicate an initial banked month and year for " & All_SNAP_Clients_Array (1, clt_err_check)
				IF  All_SNAP_Clients_Array (8, clt_err_check)  = "" THEN _
				err_msg = err_msg & vbCr & "You must indicate the number of banked months approved for " & All_SNAP_Clients_Array (1, clt_err_check)
			End If
		Next
		IF clients_with_banked_mo = 0 THEN err_msg = err_msg & vbCr & "You have not indicated any clients using banked months." & _
		  vbCr & "Though you previously marked Banked Months were approved."
		IF err_msg <> "" THEN MsgBox err_msg
	Loop until err_msg = ""

	ReDim BM_Clients_Array (3, 0)

	For clt_dialog_response = 0 to (clients_on_case - 1)		'Creates an array of all the clients the worker selected as using banked months
		IF All_SNAP_Clients_Array (5, clt_dialog_response) = checked THEN
			ReDim Preserve BM_Clients_Array (3, clt_banked_mo_apprvd)
			BM_Clients_Array (0, clt_banked_mo_apprvd) = All_SNAP_Clients_Array (0, clt_dialog_response)	'Client Ref Numb
			BM_Clients_Array (1, clt_banked_mo_apprvd) = All_SNAP_Clients_Array (1, clt_dialog_response)	'Client Name
			BM_Clients_Array (3, clt_banked_mo_apprvd) = All_SNAP_Clients_Array (8, clt_dialog_response)	'Number of Banked Months Approved
			For m = 0 to (All_SNAP_Clients_Array (8, clt_dialog_response) - 1)
				IF All_SNAP_Clients_Array(6, clt_dialog_response) + m = 13 THEN
					IF BM_Clients_Array (2, clt_banked_mo_apprvd) = "" Then
						BM_Clients_Array (2, clt_banked_mo_apprvd) = 01 & "/" & (All_SNAP_Clients_Array(7, clt_dialog_response) + 1)
					Else
						BM_Clients_Array (2, clt_banked_mo_apprvd) = BM_Clients_Array (2, clt_banked_mo_apprvd) & " & " & 01 & "/" & (All_SNAP_Clients_Array(7, clt_dialog_response) + 1)
					End IF
				ElseIf All_SNAP_Clients_Array(6, clt_dialog_response) + m = 14 THEN
					IF BM_Clients_Array (2, clt_banked_mo_apprvd) = "" Then
						BM_Clients_Array (2, clt_banked_mo_apprvd) = 02 & "/" & (All_SNAP_Clients_Array(7, clt_dialog_response) + 1)
					Else
						BM_Clients_Array (2, clt_banked_mo_apprvd) = BM_Clients_Array (2, clt_banked_mo_apprvd) & " & " & 02 & "/" & (All_SNAP_Clients_Array(7, clt_dialog_response) + 1)
					End IF
				Else
					IF BM_Clients_Array (2, clt_banked_mo_apprvd) = "" Then
						BM_Clients_Array (2, clt_banked_mo_apprvd) = (All_SNAP_Clients_Array(6, clt_dialog_response) + m) & "/" & All_SNAP_Clients_Array(7, clt_dialog_response)
					Else
						BM_Clients_Array (2, clt_banked_mo_apprvd) = BM_Clients_Array (2, clt_banked_mo_apprvd) & " & " & (All_SNAP_Clients_Array(6, clt_dialog_response) + m) & "/" & All_SNAP_Clients_Array(7, clt_dialog_response)
					End IF
				End If
			Next

			Used_ABAWD_Months_Array = Split (BM_Clients_Array (2, clt_banked_mo_apprvd), "&")	'Creates an array of all BANKED MONTHS approved

			'This for... loop will write a TIKL to review SNAP and update tracking for each banked month
			For each month_to_tikl in Used_ABAWD_Months_Array

				month_to_tikl = Trim(month_to_tikl)
				IF len(month_to_tikl) = 4 THEN month_to_tikl = "0" & month_to_tikl
				tikl_month_mm = left(month_to_tikl,2)
				tikl_month_yy = right(month_to_tikl,2)
				tikl_date = tikl_month_mm & "/01/20" & tikl_month_yy
				IF cdate(tikl_date) > date THEN 'We can only enter TIKL's for tomorrow or later
					navigate_to_MAXIS_screen "DAIL", "WRIT"
					EMWriteScreen tikl_month_mm, 5, 18
					EMWriteScreen "01", 5, 21
					EMWriteScreen tikl_month_yy, 5, 24
					transmit
					EMReadScreen tikl_corr, 4, 24, 2
					IF tikl_corr = "DATE" then
						PF10
						PF3
						tikl_set = False
						MsgBox "*** ALERT !!! ***" & vbCr & "The TIKL to review SNAP BANKED MONTHS was not set for some reason!" & vbCr & "You must set a TIKL Manually"
					Else
						tikl_notc = "BANKED MONTH CASE Must be reviewed.  Review eligibility,"
						tikl_notc_two = "update WREG tracking and appove results."
						EMWriteScreen tikl_notc, 9, 3
						EMWriteScreen tikl_notc_two, 10, 3
						tikl_set = TRUE
						cls_month = abs(last_month_mm) + 1
						IF cls_month = 13 then
							cls_date = "01" & "/" & (abs(last_month_yy) + 1)
						Else
							cls_date = cls_month & "/" & last_month_yy
						End IF
					END IF
				IF tikl_set = FALSE THEN MsgBox "*** ATTENTION ***" & vbCr & "The TIKL to review Banked Months did not set" & vbCr & "You must manually set the TIKL!"
			End If

		NEXT
		clt_banked_mo_apprvd = clt_banked_mo_apprvd + 1
		END IF
	Next
End IF

all_elig_results = 0

'Gathers all programs with elig results from ELIG SUMM and adds them to an array
'The array is per elig program and month
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
				Redim Preserve bene_amount_array (10, all_elig_results)
				EMReadScreen prog_to_check, 4, row, 22
				'EMReadScreen snap_month, 2, 19, 56
				'EMReadScreen snap_year, 2, 19, 59
				prog_to_check = trim(prog_to_check)
				bene_amount_array (0, all_elig_results) = prog_to_check
				bene_amount_array (1, all_elig_results) = cur_month
				bene_amount_array (2, all_elig_results) = cur_year
				all_elig_results = all_elig_results + 1
			End If
		End If
	Next
Next

infant_on_case = "Unknown"

'Here the script will use the program listed in the array to determine where to go to find the amounts - then add them to the array
For all_elig_results = 0 to UBound (bene_amount_array,2)
	If bene_amount_array(0, all_elig_results) = "Food" AND snap_approved_check = checked Then
		Call navigate_to_MAXIS_screen("ELIG", "FS")
		EMWriteScreen bene_amount_array (1, all_elig_results), 19, 54
		EMWriteScreen bene_amount_array (2, all_elig_results), 19, 57
		EMWRiteScreen "FSSM", 19, 70
		transmit
		EMReadScreen notc_type, 8, 3, 3
		If notc_type = "APPROVED" then
			EMReadScreen snap_bene_amt, 8, 13, 73
			EMReadScreen snap_reporter, 10, 8, 31
			EMReadScreen partial_bene, 8, 9, 44
			If partial_bene = "Prorated" then
				EMReadScreen prorated_date, 8, 9, 58
				bene_amount_array (4, all_elig_results) = prorated_date
			End If
			bene_amount_array (3, all_elig_results) = trim(snap_bene_amt)
			bene_amount_array (10, all_elig_results) = snap_reporter & " Reporter"
		ELSE
			EMReadScreen approval_versions, 2, 2, 18
			If trim(approval_versions) = "1" THEN
				bene_amount_array(0, all_elig_results) = "NONE"
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
						EMReadScreen prorated_date, 8, 9, 58
						bene_amount_array (4, all_elig_results) = prorated_date
					End If
					bene_amount_array (3, all_elig_results) = trim(snap_bene_amt)
					bene_amount_array (10, all_elig_results) = trim(snap_reporter) & " Reporter"
				Else
					IF snap_approved_check = checked THEN MsgBox "Your most recent SNAP approval for the benefit month chosen is not from today. This approval amount will not be case noted"
					bene_amount_array(0, all_elig_results) = "NONE"
				End If
			End If
		End If
	ElseIf bene_amount_array(0, all_elig_results) = "MFIP" AND cash_approved_check = checked Then
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
		EMWriteScreen bene_amount_array(1, all_elig_results), 20, 56
		EMWriteScreen bene_amount_array(2, all_elig_results), 20, 59
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
				bene_amount_array (5, all_elig_results) = trim(mfip_bene_cash_amt)
				bene_amount_array (3, all_elig_results) = trim(mfip_bene_food_amt)
				bene_amount_array (6, all_elig_results) = trim(mfip_bene_housing_amt)
				bene_amount_array (10, all_elig_results) = trim(mfip_reporter) & " Reporter"
				If prorate_date <> "        " Then bene_amount_array (4, all_elig_results) = prorate_date
			Else
				IF cash_approved_check = checked THEN MsgBox "This MFIP approval was not done today and the benefit amount will not be case noted"
				bene_amount_array(0, all_elig_results) = "NONE"
			End If
		Else
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of CASH in the selected benefit month. Please approve before running the script."
				bene_amount_array(0, all_elig_results) = "NONE"
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
					bene_amount_array (5, all_elig_results) = trim(mfip_bene_cash_amt)
					bene_amount_array (3, all_elig_results) = trim(mfip_bene_food_amt)
					bene_amount_array (6, all_elig_results) = trim(mfip_bene_housing_amt)
					bene_amount_array (10, all_elig_results) = trim(mfip_reporter) & " Reporter"
					If prorate_date <> "        " Then bene_amount_array (4, all_elig_results) = prorate_date
				Else
					IF cash_approved_check = checked THEN MsgBox "Your most recent MFIP approval is not from today and benefit amounts will not be added to case note"
					bene_amount_array(0, all_elig_results) = "NONE"
				End If
			End If
		End If
	ElseIf bene_amount_array(0, all_elig_results) = "DWP" AND cash_approved_check = checked THEN
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
		EMWriteScreen bene_amount_array(1, all_elig_results), 20, 56
		EMWriteScreen bene_amount_array(2, all_elig_results), 20, 59
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
				bene_amount_array (7, all_elig_results) = trim(DWP_bene_shel_amt)
				bene_amount_array (8, all_elig_results) = trim(DWP_bene_pers_amt)
				IF prorate_date <> "__ __ __" Then bene_amount_array (10, all_elig_results) = Replace(prorate_date, " ", "/")
			Else
				IF cash_approved_check = checked THEN MsgBox "This DWP approval was not done today and the benefit amount will not be case noted"
				bene_amount_array(0, all_elig_results) = "NONE"
			End If
		Else
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of CASH in the selected benefit month. Please approve before running the script."
				bene_amount_array(0, all_elig_results) = "NONE"
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
					bene_amount_array (7, all_elig_results) = trim(DWP_bene_shel_amt)
					bene_amount_array (8, all_elig_results) = trim(DWP_bene_pers_amt)
					IF prorate_date <> "__ __ __" Then bene_amount_array (10, all_elig_results) = Replace(prorate_date, " ", "/")
				Else
					IF cash_approved_check = checked THEN MsgBox "Your most recent DWP approval is not from today and benefit amounts will not be added to case note"
					bene_amount_array(0, all_elig_results) = "NONE"
				End If
			End If
		End If
	ElseIf bene_amount_array(0, all_elig_results) = "GA" AND cash_approved_check = checked THEN
		'GA portion
		call navigate_to_MAXIS_screen("ELIG", "GA")
		EMWriteScreen bene_amount_array(1, all_elig_results), 20, 54
		EMWriteScreen bene_amount_array(2, all_elig_results), 20, 57
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
				bene_amount_array (9, all_elig_results) = trim(GA_bene_cash_amt)
				IF prorate_date <> "     " Then bene_amount_array(10, all_elig_results) = Replace(prorate_date, " ", "/") & "/" & bene_amount_array(2,all_elig_results)
			Else
				IF cash_approved_check = checked THEN MsgBox "The most recent approval is not from today and will not be added to the case note"
				bene_amount_array(0, all_elig_results) = "NONE"
			END IF
		ELSE
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of GA in the selected benefit month. This will not be added to the case note."
				bene_amount_array(0, all_elig_results) = "NONE"
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
					bene_amount_array (9, all_elig_results) = trim(GA_bene_cash_amt)
					IF prorate_date <> "     " Then bene_amount_array(10, all_elig_results) = Replace(prorate_date, " ", "/") & "/" & bene_amount_array(2,all_elig_results)
				Else
					IF cash_approved_check = checked THEN MsgBox "The most recent approval is not from today and will not be added to the case note"
					bene_amount_array(0, all_elig_results) = "NONE"
				END IF
			End If
		END IF
	ELSEIF bene_amount_array(0, all_elig_results) = "MSA" AND cash_approved_check = checked THEN
		'MSA portion
		call navigate_to_MAXIS_screen("ELIG", "MSA")
		EMWriteScreen bene_amount_array(1, all_elig_results), 20, 56
		EMWriteScreen bene_amount_array(2, all_elig_results), 20, 59
		EMWRiteScreen "MSSM", 20, 71
		transmit
		EMReadScreen cash_approved_version, 8, 3, 3
		IF cash_approved_version = "APPROVED" THEN
			EMReadScreen cash_approval_date, 8, 3, 14
			IF cdate(cash_approval_date) = date THEN
				EMReadScreen MSA_bene_cash_amt, 8, 17, 73
				'MSA does not have proration
				bene_amount_array (9, all_elig_results) = trim(MSA_bene_cash_amt)
			Else
				IF cash_approved_check = checked THEN MsgBox "The most recent approval is not from today and will not be added to the case note"
				bene_amount_array(0, all_elig_results) = "NONE"
			END IF
		ELSE
			EMReadScreen cash_approval_versions, 1, 2, 18
			IF cash_approval_versions = "1" THEN
				IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of MSA in the selected benefit month. This will not be added to the case note"
				bene_amount_array(0, all_elig_results) = "NONE"
			Else
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 78
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				IF cdate(cash_approval_date) = date THEN
					EMReadScreen MSA_bene_cash_amt, 8, 17, 73
					'MSA does not have proration
					bene_amount_array (9, all_elig_results) = trim(MSA_bene_cash_amt)
				Else
					IF cash_approved_check = checked THEN MsgBox "You do not have an approved version of MSA in the selected benefit month. This will not be added to the case note"
					bene_amount_array(0, all_elig_results) = "NONE"
				END IF
			End If
		END IF
	END IF
Next

'Case notes----------------------------------------------------------------------------------------------------
IF clt_banked_mo_apprvd <> 0 THEN 	'BANKED MONTH Case Note - each client gets a seperate case note
	For clt_banked_mo_apprvd = 0 to UBound (BM_Clients_Array,2)
		Call start_a_blank_CASE_NOTE
		Call write_variable_in_CASE_NOTE ("!~!~! Memb " & BM_Clients_Array(0,clt_banked_mo_apprvd) & "used BANKED MONTHS in " & BM_Clients_Array (2, clt_banked_mo_apprvd) & " !~!~!")	'Months used moved to header
		IF tikl_set = TRUE Then Call write_variable_in_CASE_NOTE ("TIKL Created to review results and update tracking for each month noted above.")
		Call write_variable_in_CASE_NOTE ("---")
		Call write_variable_in_CASE_NOTE (worker_signature)
	Next
End If

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
END IF

IF hc_approved_check = checked THEN approved_programs = approved_programs & "HC/"
IF cash_approved_check = checked THEN approved_programs = approved_programs & "CASH/"
IF emer_approved_check = checked THEN approved_programs = approved_programs & "EMER/"
EMSendKey "---Approved " & approved_programs & "<backspace>" & " " & type_of_approval & "---" & "<newline>"
IF postponed_verif_check = checked THEN write_variable_in_CASE_NOTE("**EXPEDITED SNAP APPROVED BUT CASE HAS POSTPONED VERIFICATIONS.**")
IF benefit_breakdown <> "" THEN call write_bullet_and_variable_in_case_note("Benefit Breakdown", benefit_breakdown)
IF autofill_check = checked THEN
	FOR snap_approvals = 0 to UBound(bene_amount_array,2)
		IF bene_amount_array (0,snap_approvals) = "Food" AND snap_approved_check = checked THEN
			snap_header = ("SNAP for " & bene_amount_array(1,snap_approvals) & "/" & bene_amount_array(2,snap_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE (snap_header, FormatCurrency(bene_amount_array(3,snap_approvals)) & " " & bene_amount_array(10,snap_approvals))
			IF bene_amount_array (4, snap_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", bene_amount_array(4,snap_approvals))
			End If
		End If
	Next
	FOR mfip_approvals = 0 to UBound(bene_amount_array,2)
		IF bene_amount_array (0,mfip_approvals) = "MFIP" AND cash_approved_check = checked THEN
			Call write_variable_in_CASE_NOTE ("MFIP for " & bene_amount_array(1,mfip_approvals) & "/" & bene_amount_array(2,mfip_approvals) & " " & bene_amount_array(10,mfip_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE ("Cash Portion", FormatCurrency(bene_amount_array(5, mfip_approvals)))
			Call write_bullet_and_variable_in_CASE_NOTE ("Food Portion", FormatCurrency(bene_amount_array(3, mfip_approvals)))
			Call write_bullet_and_variable_in_CASE_NOTE ("Housing Grant Amount", FormatCurrency(bene_amount_array(6, mfip_approvals)))
			IF bene_amount_array (4, mfip_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", bene_amount_array(4,mfip_approvals))
			End If
		End If
	Next
	FOR dwp_approvals = 0 to UBound(bene_amount_array,2)
		IF bene_amount_array (0,dwp_approvals) = "DWP" AND cash_approved_check = checked THEN
			Call write_variable_in_CASE_NOTE ("DWP for " & bene_amount_array(1,dwp_approvals) & "/" & bene_amount_array(2,dwp_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE ("Shelter Benefit", FormatCurrency(bene_amount_array(7, dwp_approvals)))
			Call write_bullet_and_variable_in_CASE_NOTE ("Personal Needs", FormatCurrency(bene_amount_array(8, dwp_approvals)))
			IF bene_amount_array (4, dwp_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", bene_amount_array(4,dwp_approvals))
			End If
		End If
	Next
	FOR msa_approvals = 0 to UBound(bene_amount_array, 2)
		IF bene_amount_array (0,msa_approvals) = "MSA" AND cash_approved_check = checked THEN
			msa_header = ("MSA for " & bene_amount_array(1,msa_approvals) & "/" & bene_amount_array(2, msa_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE (msa_header, FormatCurrency(bene_amount_array(9,msa_approvals)))
		End If
	Next
	FOR ga_approvals = 0 to UBound(bene_amount_array, 2)
		IF bene_amount_array (0,ga_approvals) = "GA" AND cash_approved_check = checked THEN
			ga_header = ("GA for " & bene_amount_array(1,ga_approvals) & "/" & bene_amount_array(2,ga_approvals))
			Call write_bullet_and_variable_in_CASE_NOTE (ga_header, FormatCurrency(bene_amount_array(9,ga_approvals)))
			IF bene_amount_array (4, ga_approvals) <> "" THEN
				Call write_bullet_and_variable_in_CASE_NOTE ("    Prorated from: ", bene_amount_array(4,ga_approvals))
			End If
		End If
	Next
END IF
IF FIAT_checkbox = 1 THEN Call write_variable_in_CASE_NOTE ("* This case has been FIATed.")
IF other_notes <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Approval Notes", other_notes)
IF programs_pending <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Programs Pending", programs_pending)
If docs_needed <> "" then call write_bullet_and_variable_in_CASE_NOTE("Verifs needed", docs_needed)
IF SNAP_banked_mo_check = checked THEN Call write_variable_in_CASE_NOTE ("BANKED MONTHS were approved - see previous case note for detail.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)


script_end_procedure("Success! Please remember to check the generated notice to make sure it is correct. If not, please add WCOMs to make notice read correctly.")