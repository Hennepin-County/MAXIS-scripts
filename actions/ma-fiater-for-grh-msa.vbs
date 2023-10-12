'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - MA FIATER FOR GRH MSA.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 480                     'manual run time in seconds
STATS_denomination = "I"                   'C is for each CASE
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
call changelog_update("04/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

' PURPOSE: This script allows users to FIAT the HC determination for clients that are active on GRH or MSA...policy change means that these clients are no
'			longer automatically eligible for MA eligibility.
' DESIGN...
'		1. The script must collect the MAXIS Case Number and select the individual to FIAT.
'		2. The script must collect income information for the individual...
'			2a. GROSS UNEARNED INCOME
'			2b. GROSS DEEMED UNEARNED INCOME
'			2c. EXCLUDED UNEARNED INCOME
'		3. The script must collect asset information for the individual, determine which is COUNTED, EXCLUDED, and UNAVAILABLE.
'   4. More stuff tbd

'these variables are needed to input the values of each individual amount to the ELIG/HC FIAT
DIM ttl_CASH_counted, ttl_CASH_excluded, ttl_CASH_unavail
DIM ttl_ACCT_counted, ttl_ACCT_excluded, ttl_ACCT_unavail
DIM ttl_SECU_counted, ttl_SECU_excluded, ttl_SECU_unavail
DIM ttl_CARS_counted, ttl_CARS_excluded, ttl_CARS_unavail
DIM ttl_REST_counted, ttl_REST_excluded, ttl_REST_unavail
DIM ttl_OTHR_counted, ttl_OTHR_excluded, ttl_OTHR_unavail
DIM ttl_BURY_counted, ttl_BURY_excluded, ttl_BURY_unavail
DIM ttl_SPON_counted, ttl_SPON_excluded, ttl_SPON_unavail

'these variables are needed to input the values for each amount to the budget in ELIG/HC FIAT
DIM ttl_unearned_amt, ttl_earned_amt, ttl_unearned_deemed, ttl_earned_deemed

'these variables are needed for to input values to the FIAT of income

'this class is needed for to keep track of data for individual assets
class asset_object
	'variables...going to keep them public for to cut down on the work needed to manipulate
	public asset_panel
	public asset_amount
	public asset_counted_amount
	public asset_excluded_amount
	public asset_unavailable_amount
	public asset_type
	public asset_amount_dialog  	' used for display in the dialog
	public asset_type_dialog 	' used for the dialog...you'll see

	'private functions for setting asset value
	public function set_counted_amount(amount)
		asset_counted_amount = amount
	end function

	public function set_excluded_amount(amount)
		asset_excluded_amount = amount
	end function

	public function set_unavailable_amount(amount)
		asset_unavailable_amount = amount
	end function

	' function to read the amount for a specific asset
	public function read_asset_amount(len, row, col)
		EMReadScreen asset_amt, len, row, col
		asset_amt = replace(asset_amt, "_", "")
		asset_amt = trim(asset_amt)
		IF asset_amt = "" THEN asset_amt = 0
		IF asset_amt < 0 THEN 																' }
			MsgBox "ERROR: Asset found with negative balance. The script will now stop."  	' }
			stopscript																		' } should probably just have object function reject negative balance
		END IF
		IF asset_type = "COUNTED" THEN
			CALL set_counted_amount(asset_amt)
			CALL set_excluded_amount(0)
			CALL set_unavailable_amount(0)
		ELSEIF asset_type = "EXCLUDED" THEN
			CALL set_counted_amount(0)
			CALL set_excluded_amount(asset_amt)
			CALL set_unavailable_amount(0)
		ELSEIF asset_type = "UNAVAILABLE" THEN
			CALL set_counted_amount(0)
			CALL set_excluded_amount(0)
			CALL set_unavailable_amount(asset_amt)
		END IF
	end function

	' function to read whether the asset is counted
	public function read_asset_counted(row, col)
		EMReadScreen asset_counted, 1, row, col
		IF asset_counted = "Y" THEN asset_type = "COUNTED"
		IF asset_counted = "N" OR asset_counted = "_" THEN asset_type = "EXCLUDED"
	end function

	' function to assign value to panel name
	public function set_asset_panel(panel_name)
		asset_panel = panel_name
	end function

	' function to re-set amount of the asset
	public function set_asset_amount(specified_amount)
		asset_amount = specified_amount
	end function

	' function to re-set whether or not the asset is counted
	public function set_asset_type(user_selection)
		asset_type = user_selection
	end function
end class

' this class is going to be used for grabbing information from UNEA, JOBS, and BUSI
class income_object
	' variables for income objects
	public income_amt				' this is used to store the value of the per-pay-check income
	public monthly_income_amt		' this is used to store the value of the monthly income
	public income_frequency			' this is used to store the values "Monthly", "Semi-Monthly"... to reflect pay frequency
	public retro_income_amt
	public prosp_income_amt
	public snap_pic_income_amt
	public grh_pic_income_amt
	public income_category			' "UNEARNED", "EARNED", "DEEMED UNEARNED", "DEEMED EARNED"
	public income_type				' type read from panel
	public income_type_code				' 2-digit code read from panel
	public income_start_date
	public income_end_date
	public budget_month
	public COLA_amount 'COLA Disregard amount read from panel'
	private pay_freq				' read from the panel...private because it is only used to calculate the monthly values
	private income_multiplier		' extrapolated from the pay_freq...private because it is only used to calculate the monthly values

	' === member functions for income object ===

	' private member function for identifying UNEA
	private sub are_we_at_unea
		row = 1																																		' }
		col = 1																																		' }
		EMSearch "(UNEA)", row, col																													' }
		IF row = 0 THEN 																															' } safeguarding that the script
			MsgBox "Invalid function application. The script cannot confirm you are on UNEA. The script will now stop.", vbCritical 				' }	finds UNEA
			stopscript																																' }
		END IF																																		' }
	end sub

	' private member function for identifying JOBS
	private sub are_we_at_jobs
		row = 1
		col = 1
		EMSearch "(JOBS)", row, col
		IF row = 0 THEN
			MsgBox "Invalid function application. The script cannot confirm you are on JOBS. The script will now stop.", vbCritical		' }	finds JOBS
			stopscript																													' }
		END IF																															' }
	end sub

	' private member function for identifying BUSI
	private sub are_we_at_busi
		row = 1
		col = 1
		EMSearch "(BUSI)", row, col
		IF row = 0 THEN
			MsgBox "Invalid function application. The script cannot confirm you are on BUSI. The script will now stop.", vbCritical		' }	finds BUSI											' }
			stopscript																													' }
		END IF																															' }
	end sub

	'this is to determine what footer month we are looking at to determine where to read panels'
	public sub check_footer_month
		panel_in_plus_one = false
		row = 1
		col = 1
		EMSearch "Month: ", row, col
		EMReadScreen panel_date, 5, row, col + 7
		panel_date = replace(panel_date, " ", "/1/")
		if panel_date > date THEN panel_in_plus_one = true
	end sub


	' private member function for calculating monthly_income_amt
	public sub calculate_monthly_income
		monthly_income_amt = income_multiplier * (income_amt * 1)
	end sub

	public function set_income_category(specific_income_category)
		income_category = specific_income_category
	end function

	public sub read_income_type
		IF income_category = "" THEN MsgBox "No category has been set for this income type. The script will not perform optimally."
		row = 1
		col = 1
		EMSearch "Inc Type: ", row, col
		IF row <> 0 THEN 			' } Then we are on the JOBS panel...
			EMReadScreen specific_income_type, 1, row, col + 10
			IF specific_income_type = "J" THEN
				specific_income_type = "WIOA"
				income_type_code = "01"
			ELSEIF specific_income_type = "W" THEN
				specific_income_type = "Wages"
				income_type_code = "02"
			ELSEIF specific_income_type = "E" THEN
				specific_income_type = "EITC"
				income_type_code = "03"
			ELSEIF specific_income_type = "G" THEN
				specific_income_type = "Experience Works"
				income_type_code = "04"
			ELSEIF specific_income_type = "F" THEN
				specific_income_type = "Federal Work Study"
				income_type_code = "05"
			ELSEIF specific_income_type = "S" THEN
				specific_income_type = "State Work Study"
				income_type_code = "06"
			ELSEIF specific_income_type = "O" THEN
				specific_income_type = "Other"
				income_type_code = "07"
			ELSEIF specific_income_type = "C" THEN
				specific_income_type = "Contract Income"
				income_type_code = "10"
			ELSEIF specific_income_type = "T" THEN
				specific_income_type = "Training Program"
				income_type_code = "16"				' <<<<< NO OTHER CORRESPONDING CODE IN FIAT/HC
			ELSEIF specific_income_type = "P" THEN
				specific_income_type = "Service Program"
				income_type_code = "16"				' <<<<< NO OTHER CORRESPONDING CODE IN FIAT/HC
			ELSEIF specific_income_type = "R" THEN
				specific_income_type = "Rehab Program"
				income_type_code = "16"				' <<<<< NO OTHER CORRESPONDING CODE IN FIAT/HC
			END IF
		ELSE						' } THEN we are on either BUSI or UNEA
			row = 1
			col = 1
			EMSearch "Income Type: ", row, col
			IF income_category = "EARNED" OR income_category = "DEEMED EARNED" THEN 			' } THEN WE ARE ON BUSI
				EMReadScreen specific_income_type, 2, row, col + 13
				IF specific_income_type = "01" THEN
					specific_income_type = "01 Farming"
					income_type_code = "11"
				ELSEIF specific_income_type = "02" THEN
					specified_income_type = "02 Real Estate"
					income_type_code = "14"
				ELSEIF specific_income_type = "03" THEN
					specific_income_type = "03 Home Product Sales"
					income_type_code = "15"
				ELSEIF specific_income_type = "04" THEN
					specific_income_type = "04 Other Sales"
					income_type_code = "16"
				ELSEIF specific_income_type = "05" THEN
					specific_income_type = "05 Personal Services"
					income_type_code = "17"
				ELSEIF specific_income_type = "06" THEN
					specific_income_type = "06 Paper Route"
					income_type_code = "18"
				ELSEIF specific_income_type = "07" THEN
					specific_income_type = "07 In-Home Daycare"
					income_type_code = "19"
				ELSEIF specific_income_type = "08" THEN
					specific_income_type = "08 Rental Income"
					income_type_code = "20"
				ELSEIF specific_income_type = "09" THEN
					specific_income_type = "09 Other"
					income_type_code = "21"
				END IF
			ELSEIF income_category = "UNEARNED"	or income_category = "DEEMED UNEARNED" THEN 		' } THEN WE ARE ON UNEA
				EMReadScreen specific_income_type, 20, row, col + 13
				income_type_code = left(specific_income_type, 2)
				IF income_type_code = "11" THEN 			' } Updating these values for when they are FIAT'd
					income_type_code = "09"				' } Because of course there are values that do not match
				ElSEIF income_type_code = "12" THEN
					income_type_code = "10"
				ELSEIF income_type_code = "13" THEN
					income_type_code = "11"
				ELSEIF income_type_code = "14" THEN
					income_type_code = "12"
				ELSEIF income_type_code = "15" THEN
					income_type_code = "13"
				ELSEIF income_type_code = "16" THEN
					income_type_code = "14"
				ELSEIF income_type_code = "17" THEN
					income_type_code = "15"
				ELSEIF income_type_code = "18" THEN
					income_type_code = "16"
				ELSEIF income_type_code = "19" THEN
					income_type_code = "17"
				ELSEIF income_type_code = "20" THEN
					income_type_code = "19"
				ELSEIF income_type_code = "22" THEN
					income_type_code = "20"
				ELSEIF income_type_code = "23" THEN
					income_type_code = "21"
				ELSEIF income_type_code = "24" THEN
					income_type_code = "22"
				ELSEIF income_type_code = "25" THEN
					income_type_code = "23"
				ELSEIF income_type_code = "26" THEN
					income_type_code = "24"
				ELSEIF income_type_code = "27" THEN
					income_type_code = "25"
				ELSEIF income_type_code = "28" THEN
					income_type_code = "26"
				ELSEIF income_type_code = "29" THEN
					income_type_code = "27"
				ELSEIF income_type_code = "30" THEN
					income_type_code = "28"
				ELSEIF income_type_code = "31" THEN
					income_type_code = "29"
				ELSEIF income_type_code = "35" THEN
					income_type_code = "08"
				ELSEIF income_type_code = "36" THEN
					income_type_code = "05"
				ELSEIF income_type_code = "37" THEN
					income_type_code = "07"
				ELSEIF income_type_code = "38" THEN
					income_type_code = "34"
				ELSEIF income_type_code = "39" THEN
					income_type_code = "36"
				ELSEIF income_type_code = "40" THEN
					income_type_code = "36"
				ELSEIF income_type_code = "44" THEN
					income_type_code = "30"
				END IF
				specific_income_type = trim(right(specific_income_type, 17))
			END IF
		END IF
		income_type = specific_income_type
	end sub

	' member functions for reading from JOBS
	public sub read_jobs_for_hc
		are_we_at_jobs
		'THis makes sure we don't count ended income
		income_ended = false
		EMReadScreen income_end_date, 8, 9, 49
		IF income_end_date <> "__ __ __" THEN
			income_end_date = cdate(replace(income_end_date, " ", "/"))
			if income_end_date < budg_month then income_ended = true
		END if


		if income_ended <> true THEN
		row = 1
		col = 1
		If budg_month >= current_plus_one THEN 'This section reads the HC income estimator for calculating future months'
			EMSearch "HC Income Estimate", row, col
			IF row = 0 THEN
				row = 1
				col = 1
				EMSearch "_ HC Est", row, col
				CALL write_value_and_transmit("X", row, col)
			END IF
			EMReadScreen hc_jobs_amount, 8, 11, 63
			hc_jobs_amount = replace(hc_jobs_amount, "_", "")
			hc_jobs_amount = trim(hc_jobs_amount)
			IF hc_jobs_amount = "" THEN hc_jobs_amount = 0.00
			transmit
		END IF

		'This section reads the actual amounts from the prospective side of JOBS
		EMReadScreen pay_date_1, 8, 12, 54 'We use the paydate to determine future amounts
		income_amt = 0.00 'reset
		For wage_row =  12 to 16
			EMReadScreen actual_gross, 8, wage_row, 67
			actual_gross = replace(actual_gross, "_", "")
			actual_gross = trim(actual_gross)
			IF actual_gross <> "" THEN income_amt = income_amt + actual_gross
			'msgbox income_amt
		NEXT
		'Now figure out the future months
		IF budg_month > current_plus_one THEN
			EMReadScreen pay_freq, 1, 18, 35
			IF pay_freq = 1 THEN paydates_in_budg_month = 1
			IF pay_freq = 2 THEN paydates_in_budg_month = 2
			IF pay_freq = 3 THEN
				paydate_to_check = cdate(pay_date_1)
				paydates_in_budg_month = 0
				DO 'this loop counts the paydates in a month
			 		paydate_to_check = dateadd("d", 14, paydate_to_check) 'add two weeks'
					if datepart("m", paydate_to_check) = datepart("m", budg_month) THEN paydates_in_budg_month = paydates_in_budg_month +1
				LOOP UNTIL paydate_to_check >= cdate(dateadd("m", 1, budg_month))
			END IF
			IF pay_freq = 4 THEN
				paydate_to_check = cdate(pay_date_1)
				paydates_in_budg_month = 0
				DO 'this loop counts the paydates in a month
					paydate_to_check = dateadd("d", 7, paydate_to_check) 'add one weeks'
					if datepart("m", paydate_to_check) = datepart("m", budg_month) THEN paydates_in_budg_month = paydates_in_budg_month +1
				LOOP UNTIL paydate_to_check >= cdate(dateadd("m", 1, budg_month))
			END IF
			income_amt = hc_jobs_amount * paydates_in_budg_month
		END IF
		END IF
		monthly_income_amt = income_amt
		'calculate_monthly_income
	end sub

	' member function for reading from BUSI
	public sub read_busi_for_hc
		are_we_at_busi
		EMReadScreen income_amt, 8, 12, 69
		income_multiplier = 1
		calculate_monthly_income
	end sub

	' member functions for reading from UNEA
	public sub read_unea_for_hc
		are_we_at_unea
		'check_footer_month 'make sure this isn't CM+1
		'row = 1
		'col = 1
		'THis makes sure we don't count ended income
		income_ended = false
		EMReadScreen income_end_date, 8, 7, 68
		IF income_end_date <> "__ __ __" THEN
			income_end_date = cdate(replace(income_end_date, " ", "/"))
			if income_end_date < budg_month then income_ended = true
		END if

		If budg_month >= current_plus_one THEN 'This section reads the HC income estimator for calculating future months'
			row = 1
			col = 1
			EMSearch "_ HC Income Estimate", row, col
			IF row <> 0 THEN CALL write_value_and_transmit("X", row, col)
			'msgbox "reading from popup"
			EMReadScreen hc_income_info, 8, 9, 65
			EMReadScreen hc_inc_est_pay_freq, 1, 10, 63
			hc_income_info = replace(hc_income_info, "_", "")
			hc_income_info = trim(hc_income_info)
			IF hc_income_info = "" THEN hc_income_info = 0.00
			transmit
		END IF
		'Read a COLA disregard amount if it exists
		EMReadScreen COLA_amount, 8, 10, 67
		IF isnumeric(COLA_amount) = false THEN COLA_amount = 0.00
		'This section reads the actual amounts from the prospective side of UNEA
		EMReadScreen pay_date_1, 8, 13, 68 'We use the paydate to determine future amounts
		income_amt = 0.00 'reset
		For unea_row =  13 to 17
			EMReadScreen actual_gross, 8, unea_row, 68
			actual_gross = replace(actual_gross, "_", "")
			actual_gross = trim(actual_gross)
			IF actual_gross <> "" THEN income_amt = income_amt + actual_gross
			'msgbox income_amt
		NEXT
		'Now figure out the future months
		IF budg_month > current_plus_one THEN
			If hc_inc_est_pay_freq = "_" Then
				EMReadScreen panel_type, 25, 2, 28
				EMReadScreen panel_memb, 2, 4, 33
				EMReadScreen panel_instance, 1, 2, 73
				script_run_lowdown = script_run_lowdown & vbCr & "PANEL Information: " & trim(panel_type) & " " & panel_memb & " 0" & panel_instance
				script_end_procedure_with_error_report("The MA FIATer for GRH/MSA cannot run because the HC Income Estimate pop-up is not completed and the pay frequency is not entered. The HC Income Estimate is required for all income panels on a case with HC.")
			End If
			IF hc_inc_est_pay_freq = 1 THEN paydates_in_budg_month = 1
			IF hc_inc_est_pay_freq = 2 THEN paydates_in_budg_month = 2
			IF hc_inc_est_pay_freq = 3 THEN
				paydate_to_check = cdate(pay_date_1)
				paydates_in_budg_month = 0
				DO 'this loop counts the paydates in a month
					paydate_to_check = dateadd("d", 14, paydate_to_check) 'add two weeks'
					if datepart("m", paydate_to_check) = datepart("m", budg_month) THEN paydates_in_budg_month = paydates_in_budg_month +1
				LOOP UNTIL paydate_to_check >= cdate(dateadd("m", 1, budg_month))
			END IF
			IF hc_inc_est_pay_freq = 4 THEN
				paydate_to_check = cdate(pay_date_1)
				paydates_in_budg_month = 0
				DO 'this loop counts the paydates in a month
					paydate_to_check = dateadd("d", 7, paydate_to_check) 'add one weeks'
					if datepart("m", paydate_to_check) = datepart("m", budg_month) THEN paydates_in_budg_month = paydates_in_budg_month +1
				LOOP UNTIL paydate_to_check >= cdate(dateadd("m", 1, budg_month))
			END IF
			income_amt = hc_income_info * paydates_in_budg_month
		END IF
		monthly_income_amt = income_amt
	end sub
end class


'FUNCTION ======================================
FUNCTION calculate_assets(input_array, asset_counted_total)
	number_of_assets = ubound(input_array)

	'parralel array for user input
	redim parallel_array(number_of_assets, 2)

	'determining height of dialog
	dialog_height = 115 + (20 * number_of_assets)

	Do
		DO
			asset_counted_total = 0
			asset_excluded_total = 0
			asset_unavailable_total = 0
			'calculating the values of the totals...
			FOR i = 0 TO number_of_assets
				If isempty(input_array(i)) = FALSE Then
					parallel_array(i, 0) = input_array(i).asset_counted_amount
					parallel_array(i, 1) = input_array(i).asset_excluded_amount
					parallel_array(i, 2) = input_array(i).asset_unavailable_amount

					IF isempty(parallel_array(i, 0)) = true THEN parallel_array(i, 0) = 0
					IF isempty(parallel_array(i, 1)) = true THEN parallel_array(i, 1) = 0
					IF isempty(parallel_array(i, 2)) = true THEN parallel_array(i, 2) = 0

					asset_counted_total = asset_counted_total + (input_array(i).asset_counted_amount * 1)
					asset_excluded_total = asset_excluded_total + (input_array(i).asset_excluded_amount * 1)
					asset_unavailable_total = asset_unavailable_total + (input_array(i).asset_unavailable_amount * 1)
				Else
					asset_counted_total = 0
					asset_excluded_total = 0
					asset_unavailable_total = 0
				End If
			NEXT

		     BeginDialog Dialog1, 0, 0, 385, dialog_height, "Asset Dialog"
		       Text 10, 10, 55, 10, "ASSET PANEL"
			   Text 75, 10, 55, 10, "COUNTED"
			   Text 130, 10, 55, 10, "EXCLUDED"
			   Text 185, 10, 55, 10, "UNAVAILABLE"
			   FOR i = 0 TO number_of_assets
			   	If isempty(input_array(i)) = FALSE Then
			     	Text 10,  25 + (i * 20), 40, 10, input_array(i).asset_panel
					EditBox 75,  20 + (i * 20), 40, 15, parallel_array(i, 0)
					EditBox 130, 20 + (i * 20), 40, 15, parallel_array(i, 1)
					EditBox 185, 20 + (i * 20), 40, 15, parallel_array(i, 2)
				Else
					Text 10,  25 + (i * 20), 200, 10, "This case has no asset panels in STAT."
				End If
		       NEXT
		       Text 10, dialog_height - 40, 60, 10, "COUNTED Total:"
		       EditBox 70, dialog_height - 45, 50, 15, asset_counted_total & ""
		       Text 130, dialog_height - 40, 60, 10, "EXCLUDED Total:"
		       EditBox 195, dialog_height - 45, 50, 15, asset_excluded_total & ""
		       Text 250, dialog_height - 40, 70, 10, "UNAVAILABLE Total:"
		       EditBox 325, dialog_height - 45, 50, 15, asset_unavailable_total & ""
		       ButtonGroup ButtonPressed
		         OkButton 10, dialog_height - 20, 50, 15
		         CancelButton 60, dialog_height - 20, 50, 15
		         If isempty(input_array(0)) = FALSE Then PushButton 320, dialog_height - 20, 55, 15, "CALCULATE", calculator_button
		     EndDialog

			DIALOG Dialog1
			cancel_confirmation
			IF ButtonPressed = calculator_button THEN
				'Changing the values of the
				FOR i = 0 TO number_of_assets
					if parallel_array(i, 0) = "" then parallel_array(i, 0) = 0
					if parallel_array(i, 1) = "" then parallel_array(i, 1) = 0
					if parallel_array(i, 2) = "" then parallel_array(i, 2) = 0

					CALL input_array(i).set_counted_amount(parallel_array(i, 0))
					CALL input_array(i).set_excluded_amount(parallel_array(i, 1))
					CALL input_array(i).set_unavailable_amount(parallel_array(i, 2))
				NEXT
			END IF
		LOOP UNTIL ButtonPressed = -1
		call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	'Re-Calculating the values of assets
	asset_counted_total = 0
	asset_excluded_total = 0
	asset_unavailable_total = 0
	If isempty(input_array(0)) = FALSE Then
		FOR i = 0 TO number_of_assets
			IF (parallel_array(i, 0)) = "" THEN parallel_array(i, 0) = 0
			IF (parallel_array(i, 1)) = "" THEN parallel_array(i, 1) = 0
			IF (parallel_array(i, 2)) = "" THEN parallel_array(i, 2) = 0

			asset_counted_total = asset_counted_total + parallel_array(i, 0)
			asset_excluded_total = asset_excluded_total + parallel_array(i, 1)
			asset_unavailable_total = asset_unavailable_total + parallel_array(i, 2)

			CALL input_array(i).set_counted_amount(parallel_array(i, 0))
			CALL input_array(i).set_excluded_amount(parallel_array(i, 1))
			CALL input_array(i).set_unavailable_amount(parallel_array(i, 2))
		NEXT
	End If

END FUNCTION

FUNCTION calculate_income(input_array)
	number_of_incomes = ubound(input_array)

	number_client_incomes = 0
	number_deemed_incomes = 0

	FOR i = 0 TO number_of_incomes
		If IsObject(income_array(i)) = True Then
			IF InStr(input_array(i).income_category, "DEEMED") = 0 THEN
				number_client_incomes = number_client_incomes + 1
			ELSEIF InStr(input_array(i).income_category, "DEEMED") <> 0 THEN
				number_deemed_incomes = number_deemed_incomes + 1
			END IF
		END IF
	NEXT

	'dynamically determining the height of the monthly income dialog
	height_multiplier = 0
	IF number_client_incomes >= number_deemed_incomes THEN
		height_multiplier = number_client_incomes
	ELSEIF number_deemed_incomes > number_client_incomes THEN
		height_multiplier = number_deemed_incomes
	END IF

	FOR i = 0 TO number_of_incomes
		If IsObject(income_array(i)) = True Then
			IF InStr(input_array(i).income_category, "DEEMED") = 0 THEN
				deemed_income_exists = TRUE
			ELSEIF InStr(input_array(i).income_category, "DEEMED") <> 0 THEN
				non_deemed_income_exists = TRUE
			END IF
		END IF
	NEXT

	dlg_height = 55 + (20 * height_multiplier)
	If height_multiplier = 0 Then dlg_height = dlg_height + 20

	grp_hgt = (20 + (number_client_incomes * 20))
	If grp_hgt = 20 then grp_hgt = 45

	If deemed_income_exists = TRUE Then dlg_width = 460
	dlg_width = 250
    BeginDialog Dialog1, 0, 0, dlg_width, dlg_height, "Monthly Income"
	  client_incomes_row = 25
	  deemed_incomes_row = 25
	  FOR i = 0 TO number_of_incomes
		If IsObject(income_array(i)) = True Then
			IF InStr(input_array(i).income_category, "DEEMED") = 0 THEN
				Text 15, client_incomes_row, 45, 10, "Income Type:"
				Text 60, client_incomes_row, 40, 10, input_array(i).income_category
				Text 105, client_incomes_row, 50, 10, input_array(i).income_type
				Text 160, client_incomes_row, 40, 10, FormatCurrency(input_array(i).monthly_income_amt)
				Text 190, client_incomes_row, 40, 10, input_array(i).budget_month
				client_incomes_row = client_incomes_row + 20
			ELSEIF InStr(input_array(i).income_category, "DEEMED") <> 0 THEN
				Text 225, deemed_incomes_row, 45, 10, "Income Type:"
				Text 275, deemed_incomes_row, 75, 10, input_array(i).income_category
				Text 355, deemed_incomes_row, 60, 10, input_array(i).income_type
				Text 420, deemed_incomes_row, 40, 10, FormatCurrency(input_array(i).monthly_income_amt)
				deemed_incomes_row = deemed_incomes_row + 20
			END IF
		END IF
	  NEXT
	  If height_multiplier = 0 Then Text 15, 25, 200, 10, "No Income counted for MEMB " & hc_memb & "."
      ButtonGroup ButtonPressed
        OkButton dlg_width - 110, (dlg_height - 20), 50, 15
        CancelButton dlg_width - 60, (dlg_height - 20), 50, 15
      GroupBox 5, 5, 240, grp_hgt, "Client Income"
      IF number_deemed_incomes <> 0 THEN GroupBox 220, 5, 240, (20 + (number_deemed_incomes * 20)), "Deemed Income"
    EndDialog

	Do
		DIALOG Dialog1
		cancel_confirmation
		call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
END FUNCTION

' DIALOGS
' BeginDialog Dialog1, 0, 0, 171, 75, "Enter Case Number"
'   EditBox 95, 10, 70, 15, maxis_case_number
'   EditBox 95, 30, 20, 15, maxis_footer_month
'   EditBox 125, 30, 20, 15, maxis_footer_year
'   ButtonGroup ButtonPressed
'     OkButton 65, 55, 50, 15
'     CancelButton 115, 55, 50, 15
'   Text 10, 15, 75, 10, "MAXIS Case Number"
'   Text 10, 30, 75, 20, "Initial footer month of HC span:"
' EndDialog

'Dialog for testing'
BeginDialog Dialog1, 0, 0, 171, 95, "Enter Case Number"
  EditBox 95, 30, 70, 15, maxis_case_number
  EditBox 95, 50, 20, 15, maxis_footer_month
  EditBox 125, 50, 20, 15, maxis_footer_year
  ButtonGroup ButtonPressed
    OkButton 65, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 5, 5, 125, 20, "THIS SCRIPT IS IN TESTING.       Please alert the BZST to any issues."
  Text 10, 35, 75, 10, "MAXIS Case Number"
  Text 10, 50, 75, 20, "Initial footer month of HC span:"
EndDialog

testing_run = TRUE
' ================ the script ====================
EMConnect ""

CALL check_for_MAXIS(true)		' checking for MAXIS

Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(maxis_footer_month, maxis_footer_year)

Do
	DO																							' }
		err_msg = ""																			' }
																								' }
		DIALOG Dialog1																			' }
		cancel_confirmation																		' }
																								' }
		Call validate_MAXIS_case_number(err_msg, "*")											' }	initial dialog
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")	' }
																								' }
		If err_msg <> "" Then MsgBox "*****  ERROR IN DIALOG ENTRY  *****" & vbNewLine & vbNewLine & "Please resolved the following to continue:" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""																		' }
	call check_for_password(are_we_passworded_out)												' }
Loop until are_we_passworded_out = FALSE														' }

Call back_to_SELF

DO
	' Getting the individual on the case
	CALL HH_member_custom_dialog(HH_member_array)
	IF ubound(HH_member_array) <> 0 THEN MsgBox "Please pick one and only one person for this."
LOOP UNTIL ubound(HH_member_array) = 0

FOR EACH person in HH_member_array
	hc_memb = left(person, 2)
	EXIT FOR
NEXT

'Check for 1619 / waiver status, and stop if found
Call navigate_to_MAXIS_screen("STAT", "DISA")
EMWriteScreen hc_memb, 20, 76
transmit
EMReadScreen status_1619, 1, 16, 59
EMReadScreen waiver_type, 1, 14, 59
IF status_1619 <> "_" or waiver_type = "K" or waiver_type = "J" THEN script_end_procedure("This case has 1619 status or an Elderly Waiver, and the FIATer should not be used.  The script will now stop.")


' ==============
' ... ASSETS ...
' ==============
' VARIABLES WE NEED FOR THIS BIT...
'		asset_acct_amt X
'		asset_cash_amt X
'		asset_secu_amt X
'		asset_cars_amt X
'		asset_rest_amt X
'		asset_othr_amt X
'		asset_bury_amt
'		asset_spon_amt

' ==================
' ... ACCT PANEL ...
' ==================
num_assets = -1
redim asset_array(0)

'asset_acct_amt = 0													' }
CALL navigate_to_MAXIS_screen("STAT", "ACCT")						' }
EMWriteScreen hc_memb, 20, 76										' }
CALL write_value_and_transmit("01", 20, 79)							' }
EMReadScreen num_acct, 1, 2, 78										' }
IF num_acct <> "0" THEN 											' }
	Do																' }
		num_assets = num_assets + 1									' }
		redim preserve asset_array(num_assets)						' } STAT/ACCT
		set asset_array(num_assets) = new asset_object				' }
		asset_array(num_assets).set_asset_panel "ACCT"				' }
		asset_array(num_assets).read_asset_counted 14, 64			' }
		asset_array(num_assets).read_asset_amount 8, 10, 46			' }
		transmit													' }
		EMReadScreen enter_a_valid, 21, 24, 2						' }
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO		' }
	LOOP															' }
END IF

' ==================
' ... CASH PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "CASH")						' }
CALL write_value_and_transmit(hc_memb, 20, 76)						' }
EMReadScreen number_of_cash, 1, 2, 78								' }
IF number_of_cash <> "0" THEN 										' }
	num_assets = num_assets + 1										' }
	redim preserve asset_array(num_assets)							' }
	set asset_array(num_assets) = new asset_object					' } STAT/CASH
	asset_array(num_assets).set_asset_panel "CASH"					' }
	asset_array(num_assets).set_asset_type "COUNTED"				' }
	asset_array(num_assets).read_asset_amount 8, 8, 39				' }
END IF																' }

' ==================
' ... OTHR PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "OTHR")							' }
EMWriteScreen hc_memb, 20, 76											' }
CALL write_value_and_transmit("01", 20, 79)								' }
EMReadScreen number_of_other, 1, 2, 78									' }
IF number_of_other <> "0" THEN 											' }
	DO																	' }
		num_assets = num_assets + 1										' }
		redim preserve asset_array(num_assets)							' }
		set asset_array(num_assets) = new asset_object					' } STAT/OTHR
		asset_array(num_assets).set_asset_panel "OTHR"					' }
		asset_array(num_assets).read_asset_counted 12, 64				' }
		asset_array(num_assets).read_asset_amount 10, 8, 40				' }
		transmit														' }
		EMReadScreen enter_a_valid, 21, 24, 2							' }
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO			' }
	LOOP																' }
END IF																	' }

' ==================
' ... SECU PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "SECU")							' }
EMWriteScreen hc_memb, 20, 76											' }
CALL write_value_and_transmit("01", 20, 79)								' }
EMReadScreen number_of_secu, 1, 2, 78									' }
IF number_of_secu <> "0" THEN 											' }
	DO																	' }
		num_assets = num_assets + 1										' }
		redim preserve asset_array(num_assets)							' }	STAT/SECU
		set asset_array(num_assets) = new asset_object					' }
		CALL asset_array(num_assets).set_asset_panel("SECU")			' }
		CALL asset_array(num_assets).read_asset_counted(15, 64)			' }
		CALL asset_array(num_assets).read_asset_amount(8, 10, 52)		' }
		transmit														' }
		EMReadScreen enter_a_valid, 21, 24, 2							' }
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO			' }
	LOOP																' }
END IF																	' }

' ==================
' ... CARS PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "CARS")							' }
EMWriteScreen hc_memb, 20, 76											' }
CALL write_value_and_transmit("01", 20, 79)								' }
EMReadScreen number_of_cars, 1, 2, 78									' }
IF number_of_cars <> "0" THEN 											' }
	DO																	' }
		num_assets = num_assets + 1										' }
		redim preserve asset_array(num_assets)							' }
		set asset_array(num_assets) = new asset_object					' } STAT/CARS
		CALL asset_array(num_assets).set_asset_amount("CARS")			' }
		CALL asset_array(num_assets).read_asset_counted(15, 76)			' }
		CALL asset_array(num_assets).read_asset_amount(8, 9, 45)		' }
		transmit														' }
		EMReadScreen enter_a_valid, 21, 24, 2							' }
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO			' }
	LOOP																' }
END IF																	' }

CALL calculate_assets(asset_array, asset_counted_total)

' creating totals for the ttl_whatever variables for to FIAT the assets
If isempty(asset_array(0)) = FALSE Then
	FOR i = 0 TO ubound(asset_array)
		IF asset_array(i).asset_panel = "ACCT" THEN
			ttl_ACCT_counted = ttl_ACCT_counted +   (1 * asset_array(i).asset_counted_amount)
			ttl_ACCT_excluded = ttl_ACCT_excluded + (1 * asset_array(i).asset_excluded_amount)
			ttl_ACCT_unavail = ttl_ACCT_unavail +   (1 * asset_array(i).asset_unavailable_amount)
		ELSEIF asset_array(i).asset_panel = "CARS" THEN
			ttl_CARS_counted = ttl_CARS_counted +   (1 * asset_array(i).asset_counted_amount)
			ttl_CARS_excluded = ttl_CARS_excluded + (1 * asset_array(i).asset_excluded_amount)
			ttl_CARS_unavail = ttl_CARS_unavail +   (1 * asset_array(i).asset_unavailable_amount)
		ELSEIF asset_array(i).asset_panel = "CASH" THEN
			ttl_CASH_counted = ttl_CASH_counted + (1 * asset_array(i).asset_counted_amount)
			ttl_CASH_excluded = ttl_CASH_excluded + (1 * asset_array(i).asset_excluded_amount)
			ttl_CASH_unavail = ttl_CASH_unavail + (1 * asset_array(i).asset_unavailable_amount)
		ELSEIF asset_array(i).asset_panel = "OTHR" THEN
			ttl_OTHR_counted = ttl_OTHR_counted + (1 * asset_array(i).asset_counted_amount)
			ttl_OTHR_excluded = ttl_OTHR_excluded + (1 * asset_array(i).asset_excluded_amount)
			ttl_OTHR_unavail = ttl_OTHR_unavail + (1 * asset_array(i).asset_unavailable_amount)
		ELSEIF asset_array(i).asset_panel = "REST" THEN
			ttl_REST_counted = ttl_REST_counted + (1 * asset_array(i).asset_counted_amount)
			ttl_REST_excluded = ttl_REST_excluded + (1 * asset_array(i).asset_excluded_amount)
			ttl_REST_unavail = ttl_REST_unavail + (1 * asset_array(i).asset_unavailable_amount)
		ELSEIF asset_array(i).asset_panel = "SECU" THEN
			ttl_SECU_counted = ttl_SECU_counted + (1 * asset_array(i).asset_counted_amount)
			ttl_SECU_excluded = ttl_SECU_excluded + (1 * asset_array(i).asset_excluded_amount)
			ttl_SECU_unavail = ttl_SECU_unavail + (1 * asset_array(i).asset_unavailable_amount)
		END IF
	NEXT
End If

CALL check_for_MAXIS(false) 	' checking for MAXIS again again


CALL check_for_MAXIS(false) 	' checking for MAXIS again again

'The business of FIATing
CALL navigate_to_MAXIS_screen("ELIG", "HC")

'finding the correct household member
FOR hhmm_row = 8 to 19
	EMReadScreen hhmm_pers, 2, hhmm_row, 3
	IF hhmm_pers = hc_memb THEN EXIT FOR
NEXT

EMReadScreen ma_case, 4, hhmm_row, 26					' }
IF ma_case <> "_ MA" THEN 								' } looking to see that the client has MA
	script_end_procedure_with_error_report("The script is not reading an MA span available for the person(s) selected on this case. If this is incorrect, send an error report to the BZ Script Team for review of the code.")
End If

CALL write_value_and_transmit("X", hhmm_row, 26)		' navigating to BSUM for that client's MA

PF9													' }
'checking if FIAT already...						' }
EMReadScreen cannot_fiat, 20, 24, 2					' }
IF cannot_fiat <> "PF9 IS NOT PERMITTED" THEN 		' }
	EMSendKey "04"									' } FIAT 500 for POLICY CHANGE
	transmit										' }
END IF												' }

'FIAT Millecento the Assets
CALL write_value_and_transmit("X", 7, 17)			' } gets to MAPT
CALL write_value_and_transmit("X", 7, 3)			' } gets to ASSETS popup


' wiping existing values...
FOR row = 10 to 17
	for col = 35 to 63 step 14
		EMWriteScreen "__________", row, col
	next
NEXT

' writing total counted, excluded, and unavailable amounts
EMWriteScreen ttl_CASH_counted, 10, 35
EMWriteScreen ttl_CASH_excluded, 10, 49
EMWriteScreen ttl_CASH_unavailable, 10, 63
EMWriteScreen ttl_ACCT_counted, 11, 35
EMWriteScreen ttl_ACCT_excluded, 11, 49
EMWriteScreen ttl_ACCT_unavailable, 11, 63
EMWriteScreen ttl_SECU_counted, 12, 35
EMWriteScreen ttl_SECU_excluded, 12, 49
EMWriteScreen ttl_SECU_unavailable, 12, 63
EMWriteScreen ttl_CARS_counted, 13, 35
EMWriteScreen ttl_CARS_excluded, 13, 49
EMWriteScreen ttl_CARS_unavailable, 13, 63
EMWriteScreen ttl_REST_counted, 14, 35
EMWriteScreen ttl_REST_excluded, 14, 49
EMWriteScreen ttl_REST_unavailable, 14, 63
EMWriteScreen ttl_OTHR_counted, 15, 35
EMWriteScreen ttl_OTHR_excluded, 15, 49
EMWriteScreen ttl_OTHR_unavailable, 15, 63


transmit
transmit
PF3

IF asset_counted_total >= 3000 THEN
	end_msg = "The client appears to exceed $3,000 in counted assets." & vbNewLine &  "Follow instructions in One Source."
	script_end_procedure(end_msg)
END IF

' ==============
' ... Income ...
' ==============
num_income = -1
redim income_array(0)

'==============
'Determmine which months to collect data from
initial_month = cdate(maxis_footer_month & "/01/20" & maxis_footer_year)
budg_month_2 = dateadd("m", initial_month, 1)
budg_month_3 = dateadd("m", initial_month, 2)
budg_month_4 = dateadd("m", initial_month, 3)
budg_month_5 = dateadd("m", initial_month, 4)
budg_month_6 = dateadd("m", initial_month, 5)
'find current month, so we can handle the future differently
current_month = datepart("m", date) & "/1/" & datepart("yyyy", date)

'We loop through the panels six times, starting with the oldest month in the budget span we wish to FIAT.'
budg_month = initial_month
current_plus_one = cdate(datepart("m", dateadd("m", 1, date)) & "/01/" & datepart("yyyy", dateadd("m", 1, date)))
For month_add = 0 to 5
	'Set the footer month for the current LOOP
budg_month = dateadd("m", month_add, initial_month)

maxis_footer_month = datepart("m", budg_month)
if len(maxis_footer_month) = 1 THEN maxis_footer_month = "0" & maxis_footer_month
maxis_footer_year = right(datepart("yyyy", budg_month), 2)
If budg_month > date THEN
	future_month = true
	maxis_footer_month = datepart("m", current_plus_one) 'For any month CM+1 or greater, we read CM+1
	if len(maxis_footer_month) = 1 THEN maxis_footer_month = "0" & maxis_footer_month
	maxis_footer_year = right(datepart("yyyy", current_plus_one), 2)
END IF
back_to_self 'This is necessary to get the footer months to work right'

' ====================
' ...earned income ...
' ====================
' ==================
' ... JOBS PANEL ...
' ==================

CALL navigate_to_MAXIS_screen("STAT", "JOBS")
EMWriteScreen hc_memb, 20, 76
CALL write_value_and_transmit("01", 20, 79)
EMReadScreen number_of_jobs, 1, 2, 78
IF number_of_jobs <> "0" THEN
	DO
		num_income = num_income + 1
		redim preserve income_array(num_income)
		set income_array(num_income) = new income_object
		CALL income_array(num_income).set_income_category("EARNED")
		income_array(num_income).read_jobs_for_hc
		income_array(num_income).read_income_type
		income_array(num_income).budget_month = budg_month
		transmit
		EMReadScreen enter_a_valid, 21, 24, 2
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO
	LOOP
END IF

' ==================
' ... BUSI PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "BUSI")
EMWriteScreen hc_memb, 20, 76
CALL write_value_and_transmit("01", 20, 79)

' =====================
' ...unearned income...
' =====================

' ==================
' ... UNEA PANEL ...
' ==================
CALL navigate_to_MAXIS_screen("STAT", "UNEA")
EMWriteScreen hc_memb, 20, 76
CALL write_value_and_transmit("01", 20, 79)
EMReadScreen number_of_unea, 1, 2, 78
IF number_of_unea <> "0" THEN
	DO
		num_income = num_income + 1
		redim preserve income_array(num_income)
		set income_array(num_income) = new income_object
		CALL income_array(num_income).set_income_category("UNEARNED")
		income_array(num_income).read_unea_for_hc
		income_array(num_income).read_income_type
		income_array(num_income).budget_month = budg_month
		transmit													' }
		EMReadScreen enter_a_valid, 21, 24, 2						' } navigating to the next UNEA
		IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO		' }
	LOOP
END IF
NEXT

' asking the user if there is income deeming on this case.
'		if the user says YES then the script asks for the household member.
'		the script checks to make sure the user did not select the same person as hc_memb
'		then the script grabs all income information from that individual

is_there_income_deeming = MsgBox ("The script has finished grabbing income information for the client." & vbNewLine & "Is there deeming income on this case?" & vbNewLine & vbTab & "Press YES to get the deemed income." & vbNewLine & vbTab & "Press NO to continue." & vbNewLine & vbTab & "Press CANCEL to stop the script.", vbYesNoCancel + vbInformation, "Does this case have deeming income to include?")
call check_for_MAXIS(false)
IF is_there_income_deeming = vbCancel THEN
	script_end_procedure("Script cancelled.")
ELSEIF is_there_income_deeming = vbYes THEN
	' grabbing the ref num of the deeming individual
	' and confirming it is not the same as the applicant
	DO
		DO
			' Getting the individual on the case
			CALL HH_member_custom_dialog(HH_member_array)
			IF ubound(HH_member_array) <> 0 THEN MsgBox "Please pick one and only one person for this."
		LOOP UNTIL ubound(HH_member_array) = 0

		FOR EACH person in HH_member_array
			deem_memb = left(person, 2)
			EXIT FOR
		NEXT

		IF hc_memb = deem_memb THEN
			MsgBox "You have selected the same household member. Pick a different household member whose income will deem.", vbExclamation
		ELSEIF hc_memb <> deem_memb THEN
			EXIT DO
		END IF
	LOOP

	' ==================
	' ... JOBS PANEL ...
	' ==================
	CALL navigate_to_MAXIS_screen("STAT", "JOBS")
	EMWriteScreen deem_memb, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen number_of_jobs, 1, 2, 78
	IF number_of_jobs <> "0" THEN
		DO
			num_income = num_income + 1
			redim preserve income_array(num_income)
			set income_array(num_income) = new income_object
			CALL income_array(num_income).set_income_category("DEEMED EARNED")
			income_array(num_income).read_jobs_for_hc
			income_array(num_income).read_income_type
			transmit
			EMReadScreen enter_a_valid, 21, 24, 2
			IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO
		LOOP
	END IF

	' ==================
	' ... BUSI PANEL ...
	' ==================
	CALL navigate_to_MAXIS_screen("STAT", "BUSI")
	EMWriteScreen deem_memb, 20, 76
	CALL write_value_and_transmit("01", 20, 79)

	' =====================
	' ...unearned income...
	' =====================

	' ==================
	' ... UNEA PANEL ...
	' ==================
	CALL navigate_to_MAXIS_screen("STAT", "UNEA")
	EMWriteScreen deem_memb, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen number_of_unea, 1, 2, 78
	IF number_of_unea <> "0" THEN
		DO
			num_income = num_income + 1
			redim preserve income_array(num_income)
			set income_array(num_income) = new income_object
			CALL income_array(num_income).set_income_category("DEEMED UNEARNED")
			income_array(num_income).read_unea_for_hc
			income_array(num_income).read_income_type
			income_array(num_income).budget_month = budg_month
			transmit													' }
			EMReadScreen enter_a_valid, 21, 24, 2						' } navigating to the next UNEA
			IF enter_a_valid = "ENTER A VALID COMMAND" THEN EXIT DO		' }
		LOOP
	END IF
END IF

' assigning values to the ttl_whatever variables for to FIAT the budget
FOR i = 0 to ubound(income_array)
	If IsObject(income_array(i)) = True Then

		IF income_array(i).income_category = "UNEARNED" 		THEN ttl_unearned_amt = ttl_unearned_amt + (income_array(i).monthly_income_amt * 1)
		IF income_array(i).income_category = "EARNED" 			THEN ttl_earned_amt = ttl_earned_amt + (income_array(i).monthly_income_amt * 1)
		IF income_array(i).income_category = "DEEMED UNEARNED" 	THEN ttl_unearned_deemed = ttl_unearned_deemed + (income_array(i).monthly_income_amt * 1)
		IF income_array(i).income_category = "DEEMED EARNED" 	THEN ttl_earned_deemed = ttl_earned_deemed + (income_array(i).monthly_income_amt * 1)
	End if
NEXT

' putting all of our income information into a lovely dialog
CALL calculate_income(income_array)

'Resetting the footer month before attempting anything else in MAXIS
maxis_footer_month = datepart("m", initial_month)'
maxis_footer_year = right(datepart("yyyy", initial_month), 2)
if len(maxis_footer_month) = 1 THEN maxis_footer_month = "0" & maxis_footer_month


' case noting information to see what we are working with
' this can be deleted when we are done
'CALL navigate_to_MAXIS_screen("CASE", "NOTE")
'PF9
'CALL write_variable_in_case_note("Testing the GRH MSA MA FIAT thingy")
'FOR i = 0 to ubound(asset_array)
'	CALL write_variable_in_case_note(asset_array(i).asset_panel & ": " & formatcurrency(asset_array(i).asset_amount) & ", " & asset_array(i).asset_type)
'NEXT
'
'FOR i = 0 to ubound(income_array)
'	CALL write_variable_in_case_note(income_array(i).income_category & ": " & formatcurrency(income_array(i).monthly_income_amt) & ", " & income_array(i).income_type)
'NEXT


CALL check_for_MAXIS(false) 	' checking for MAXIS again again

'The business of FIATing
CALL navigate_to_MAXIS_screen("ELIG", "HC")

'finding the correct household member
FOR hhmm_row = 8 to 19
	EMReadScreen hhmm_pers, 2, hhmm_row, 3
	IF hhmm_pers = hc_memb THEN EXIT FOR
NEXT

EMReadScreen ma_case, 4, hhmm_row, 26					' }
IF ma_case <> "_ MA" THEN msgbox "error"				' } looking to see that the client has MA

CALL write_value_and_transmit("X", hhmm_row, 26)		' navigating to BSUM for that client's MA

'Make sure this is the correct type of case'
EMReadScreen method_check, 55, 13, 21
method_check = replace(method_check, " ", "")
'IF method_check <> "XXXXXX" THEN script_end_procedure("This is not an auto-ma case for the entire budget period, please process manually.  The script will now stop.")

PF9													' }
'checking if FIAT already...						' }
EMReadScreen cannot_fiat, 20, 24, 2					' }
IF cannot_fiat <> "PF9 IS NOT PERMITTED" THEN 		' }
	EMSendKey "04"									' } FIAT 500 for POLICY CHANGE
	transmit										' }
END IF												' }

'This willenter the income standard and method.

FOR i = 0 to 5
	EMWriteScreen "B", 13, (21 + (i * 11))
	EMWriteScreen "E", 12, (22 + (i * 11))
NEXT

' going through and updating the budget with income and assets
FOR i = 0 TO 5
	EMWriteScreen "X", 9, (21 + (i * 11))			' pooting the X on the BUDGET field for that month in the benefit period
NEXT



transmit
'First step through the income array and look for non-deemed SSI.  If SSI is found, all income of applicant is excluded'
income_exclusion_code = "N" 'set income exclusion to N by default '
FOR goat = 0 TO ubound(income_array)
	If IsObject(income_array(goat)) = True Then
		IF income_array(goat).income_category = "UNEARNED" THEN
			IF income_array(goat).income_type_code = "03" THEN
				IF income_array(goat).monthly_income_amt > 0 THEN income_exclusion_code = "Y" 'We exclude all income if they receive SSI
			END If
		END IF
	END IF
NEXT

'This loop steps through each month of the budgets and writes in the income
For chicken = 1 to 6'

'The script now needs to go through all the income types to make sure it is putting the correct income type in the correct field...
EMWriteScreen "N", 5, 63			' WRITING "N" for PTMA
FOR i = 0 TO ubound(income_array)

	'first check which month we're budgeting
	EMReadScreen current_budg_month, 5, 6, 11
	current_budg_month = cdate(left(current_budg_month, 3) & "01/" & right(current_budg_month, 2)) 'convert to a date'
	If IsObject(income_array(i)) = True Then
		if income_array(i).budget_month = current_budg_month and income_array(i).monthly_income_amt <> 0 THEN 'only write values from the month we're in
		IF income_array(i).income_category = "UNEARNED" THEN
			CALL write_value_and_transmit("X", 8, 3)
			fiat_unea_row = 8
			DO
				EMReadScreen blank_space_for_writing, 2, fiat_unea_row, 8
				IF blank_space_for_writing = "__" THEN EXIT DO
				fiat_unea_row = fiat_unea_row + 1
			LOOP
			EMWriteScreen income_array(i).income_type_code, fiat_unea_row, 8
			EMWriteScreen income_array(i).monthly_income_amt, fiat_unea_row, 43
			EMWriteScreen income_exclusion_code, fiat_unea_row, 58
			transmit
			PF3
			'Write the COLA if appropriate'
			IF income_array(i).COLA_amount > 0 AND datepart("M", current_budg_month) < 7 THEN
				EMWriteScreen "X", 11, 3
				transmit
				EMWriteScreen income_array(i).COLA_amount, 14, 43
				transmit
				PF3
			END IF
		ELSEIF income_array(i).income_category = "EARNED" THEN
			CALL write_value_and_transmit("X", 8, 43)
			fiat_earn_row = 8
			DO
				EMReadScreen blank_space_for_writing, 2, fiat_earn_row, 8
				IF blank_space_for_writing = "__" THEN EXIT DO
				fiat_earn_row = fiat_earn_row + 1
			LOOP
			EMWriteScreen income_array(i).income_type_code, fiat_earn_row, 8
			EMWriteScreen income_array(i).monthly_income_amt, fiat_earn_row, 43
			EMWriteScreen income_exclusion_code, fiat_earn_row, 59
			transmit
			PF3
		ELSEIF income_array(i).income_category = "DEEMED EARNED" THEN
			CALL write_value_and_transmit("X", 9, 43)
			fiat_deem_earn_row = 8
			DO
				EMReadScreen blank_space_for_writing, 2, fiat_deem_earn_row, 8
				IF blank_space_for_writing = "__" THEN EXIT DO
				fiat_deem_earn_row = fiat_deem_earn_row + 1
			LOOP
			EMWriteScreen income_array(i).income_type_code, fiat_deem_earn_row, 8
			EMWriteScreen income_array(i).monthly_income_amt, fiat_deem_earn_row, 43
			EMWriteScreen "N", fiat_deem_earn_row, 59
			transmit
			PF3
		ELSEIF income_array(i).income_category = "DEEMED UNEARNED" THEN
			CALL write_value_and_transmit("X", 9, 3)
			fiat_deem_unea_row = 8
			DO
				EMReadScreen blank_space_for_writing, 2, fiat_deem_unea_row, 8
				IF blank_space_for_writing = "__" THEN EXIT DO
				fiat_deem_unea_row = fiat_deem_unea_row + 1
			LOOP
			EMWriteScreen income_array(i).income_type_code, fiat_deem_unea_row, 8
			EMWriteScreen income_array(i).monthly_income_amt, fiat_deem_unea_row, 43
			IF income_array(i).income_type_code = "03" THEN
				EMWriteScreen "Y", fiat_deem_unea_row, 58 'If this is SSI, code excluded'
			ELSE
			EMWriteScreen "N", fiat_deem_unea_row, 58
			END IF
			transmit
			PF3
			'Write the COLA if appropriate'
			IF income_array(i).COLA_amount > 0 AND datepart("M", budg_month) < 7 THEN
				EMWriteScreen "X", 11, 3
				transmit
				EMWriteScreen income_array(i).COLA_amount, 14, 43
				transmit
				PF3
			END IF
		END IF
		END IF
	END IF
NEXT
transmit

NEXT 'closing out the chicken loop'

'Now enter assets

'FIAT Millecento the Assets
For i = 1 to 6 'mark the person tests'
EMWriteScreen "X", 7, (i*11) + 6
Next
transmit

DO ' This loop goes through each available MAPT screen and enters the assets on the popup '
EMReadScreen MAPT_check, 4, 3, 51
IF MAPT_check <> "MAPT" THEN EXIT DO
CALL write_value_and_transmit("X", 7, 3)			' } gets to ASSETS popup


' wiping existing values...
FOR row = 10 to 17
	for col = 35 to 63 step 14
		EMWriteScreen "__________", row, col
	next
NEXT

' writing total counted, excluded, and unavailable amounts
EMWriteScreen ttl_CASH_counted, 10, 35
EMWriteScreen ttl_CASH_excluded, 10, 49
EMWriteScreen ttl_CASH_unavailable, 10, 63
EMWriteScreen ttl_ACCT_counted, 11, 35
EMWriteScreen ttl_ACCT_excluded, 11, 49
EMWriteScreen ttl_ACCT_unavailable, 11, 63
EMWriteScreen ttl_SECU_counted, 12, 35
EMWriteScreen ttl_SECU_excluded, 12, 49
EMWriteScreen ttl_SECU_unavailable, 12, 63
EMWriteScreen ttl_CARS_counted, 13, 35
EMWriteScreen ttl_CARS_excluded, 13, 49
EMWriteScreen ttl_CARS_unavailable, 13, 63
EMWriteScreen ttl_REST_counted, 14, 35
EMWriteScreen ttl_REST_excluded, 14, 49
EMWriteScreen ttl_REST_unavailable, 14, 63
EMWriteScreen ttl_OTHR_counted, 15, 35
EMWriteScreen ttl_OTHR_excluded, 15, 49
EMWriteScreen ttl_OTHR_unavailable, 15, 63
transmit
transmit
transmit
LOOP



'=========>>>>>>>> Here we go back to ELIG and check for potential spendown standard.



script_end_procedure_with_error_report("Success.  Please review your results before approving.")
