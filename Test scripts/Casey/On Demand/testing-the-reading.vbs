'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - REVIEW REPORT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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
call changelog_update("10/15/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'This is a script specific function and will not work outside of this script.



'constants for review_array
const worker_const          = 0
const case_number_const     = 1
const interview_const       = 2
const no_interview_const    = 3
const current_SR_const      = 4
const MFIP_status_const     = 5
const DWP_status_const      = 6
const GA_status_const       = 7
const MSA_status_const      = 8
const GRH_status_const      = 9
const CASH_next_SR_const    = 10
const CASH_next_ER_const    = 11
const SNAP_status_const     = 12
const SNAP_SR_status_const  = 13
const SNAP_next_SR_const	= 14
const SNAP_next_ER_const    = 15
const MA_status_const       = 16
const MSP_status_const      = 17
const HC_SR_status_const	= 18
const HC_ER_status_const	= 19
const HC_next_SR_const      = 20
const HC_next_ER_const      = 21
const Language_const        = 22
const Interpreter_const     = 23
const phone_1_const         = 24
const phone_2_const         = 25
const phone_3_const         = 26
const CASH_revw_status_const= 27
const SNAP_revw_status_const= 28
const HC_revw_status_const	= 29
const HC_MAGI_code_const	= 30
const review_recvd_const	= 31
const interview_date_const	= 32
const saved_to_excel_const	= 33
const notes_const           = 34

DIM review_array()              'declaring the array
ReDim review_array(notes_const, 0)       're-establihing size of array.

REPT_month = CM_plus_1_mo
REPT_year  = CM_plus_1_yr
MAXIS_footer_month = REPT_month
MAXIS_footer_year = REPT_year

report_date = REPT_month & "-" & REPT_year  'establishing review date

review_report_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"

call excel_open(review_report_file_path, True, True, ObjExcel, objWorkbook)

'==================== thE REVIEWING ========================'
' incrementor_var = 0
'
' Do
'
'     Dialog1 = ""
'     BeginDialog Dialog1, 0, 0, 126, 55, "Dialog"
'       EditBox 85, 10, 35, 15, excel_row
'       ButtonGroup ButtonPressed
'         OkButton 70, 35, 50, 15
'       Text 5, 15, 75, 10, "Select the Excel Row:"
'     EndDialog
'
'     dialog Dialog1
'     cancel_confirmation
'
'
' 	ReDim Preserve review_array(notes_const, incrementor_var)
'
' 	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
'
'
' 	Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) 'function to check PRIV status
' 	If is_this_priv = True then
' 		review_array(notes_const, incrementor_var) = "PRIV Case."
' 		review_array(interview_const, incrementor_var) = ""
' 		review_array(no_interview_const, incrementor_var) = ""
' 		review_array(current_SR_const, incrementor_var) = ""
' 	Else
' 		EmReadscreen worker_prefix, 4, 21, 14
' 		If worker_prefix <> "X127" then
' 			review_array(notes_const, i) = "Out-of-County: " & right(worker_prefix, 2)
' 			review_array(notes_const, incrementor_var) = "PRIV Case."
' 			review_array(interview_const, incrementor_var) = ""
' 			review_array(no_interview_const, incrementor_var) = ""
' 			review_array(current_SR_const, incrementor_var) = ""
' 		Else
' 			'function to determine programs and the program's status---Yay Casey!
' 			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
'
' 			If case_active = False then
' 				review_array(notes_const, incrementor_var) = "Case Not Active."
' 			Else
' 				'valuing the array variables from the inforamtion gathered in from CASE/CURR
' 				review_array(MFIP_status_const, incrementor_var) = mfip_case
' 				review_array(DWP_status_const,  incrementor_var) = dwp_case
' 				review_array(GA_status_const,   incrementor_var) = ga_case
' 				review_array(MSA_status_const,  incrementor_var) = msa_case
' 				review_array(GRH_status_const,  incrementor_var) = grh_case
' 				review_array(SNAP_status_const, incrementor_var) = snap_case
' 				review_array(MA_status_const,   incrementor_var) = ma_case
' 				review_array(MSP_status_const,  incrementor_var) = msp_case
' 				'----------------------------------------------------------------------------------------------------STAT/REVW
' 				CALL navigate_to_MAXIS_screen("STAT", "REVW")
'
' 				If family_cash_case = True or adult_cash_case = True or grh_case = True then
' 					'read the CASH review information
' 					Call write_value_and_transmit("X", 5, 35) 'CASH Review Information
' 					EmReadscreen cash_review_popup, 11, 5, 35
' 					MsgBox "HERE"
' 					If cash_review_popup = "GRH Reports" then
' 					'The script will now read the CSR MO/YR and the Recert MO/YR
' 						EMReadScreen CSR_mo, 2, 9, 26
' 						EMReadScreen CSR_yr, 2, 9, 32
' 						EMReadScreen recert_mo, 2, 9, 64
' 						EMReadScreen recert_yr, 2, 9, 70
'
' 						CASH_CSR_date = CSR_mo & "/" & CSR_yr
' 						If CASH_CSR_date = "__/__" then CASH_CSR_date = ""
'
' 						CASH_ER_date = recert_mo & "/" & recert_yr
' 						If CASH_ER_date = "__/__" then CASH_ER_date = ""
'
' 						'Comparing CSR dates to the month of REVS review
' 						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN
' 							review_array(current_SR_const, incrementor_var) = True
' 						Else
' 							review_array(current_SR_const, incrementor_var) = False
' 						End if
'
' 						'Determining if a case is ER, and if it meets interview requirement
' 						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then
' 							If mfip_case = True then review_array(interview_const, incrementor_var) = True             'MFIP interview requirement
' 							IF adult_cash_case = True or grh_case = True then review_array(no_interview_const, incrementor_var) = True    'Adult CASH programs do not meet interview requirement
' 						Elseif recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) then
' 							review_array(interview_const, incrementor_var) = False
' 							review_array(no_interview_const, incrementor_var) = False
' 						End if
'
' 						'Next CASH ER and SR dates
' 						review_array(CASH_next_SR_const, incrementor_var) = CASH_CSR_date
' 						review_array(CASH_next_ER_const, incrementor_var) = CASH_ER_date
' 					Else
' 						review_array(notes_const, incrementor_var) = "Unable to Access CASH Review Information."
' 					End if
' 					Transmit 'to exit out of the pop-up screen
' 				End if
'
' 				If snap_case = True then
' 					'read the SNAP review information
' 					Call write_value_and_transmit("X", 5, 58) 'SNAP Review Information
' 					EmReadscreen food_review_popup, 20, 5, 30
' 					If food_review_popup = "Food Support Reports" then
' 					'The script will now read the CSR MO/YR and the Recert MO/YR
' 						EMReadScreen CSR_mo, 2, 9, 26
' 						EMReadScreen CSR_yr, 2, 9, 32
' 						EMReadScreen recert_mo, 2, 9, 64
' 						EMReadScreen recert_yr, 2, 9, 70
'
' 						SNAP_CSR_date = CSR_mo & "/" & CSR_yr
' 						If SNAP_CSR_date = "__/__" then SNAP_CSR_date = ""
'
' 						SNAP_ER_date = recert_mo & "/" & recert_yr
' 						If SNAP_ER_date = "__/__" then SNAP_ER_date = ""
' 						' CSR_mo = CSR_mo & ""
' 						' CSR_yr = CSR_yr & ""
' 						' REPT_month = REPT_month & ""
' 						' REPT_year = REPT_year & ""
'
' 						'Comparing CSR and ER daates to the month of REVS review
' 						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN
' 						' IF CSR_mo = REPT_month and CSR_yr = REPT_year THEN
' 							MsgBOx "MATCH"
' 							review_array(current_SR_const, incrementor_var) = True
' 						ElseIf review_array(current_SR_const, incrementor_var) = "" Then
' 							MsgBox "No bueno"
' 							review_array(current_SR_const, incrementor_var) = False
' 						End if
'
' 						If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) then review_array(interview_const, incrementor_var) = False
' 						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(interview_const, incrementor_var) = True
'
' 						'Next SNAP ER and SR dates
' 						review_array(SNAP_next_SR_const, incrementor_var) = SNAP_CSR_date
' 						review_array(SNAP_next_ER_const, incrementor_var) = SNAP_ER_date
'
' 						' MsgBox "SR info: Panel - " & CSR_mo & "/" & CSR_yr & vbCr & "Script - " & REPT_month & "/" & REPT_year & vbCr & "BOOLEAN - " & review_array(current_SR_const, incrementor_var) & vbCr & vbCr & "ER info: Panel - " & recert_mo & "/" & recert_yr & vbCr & "Script - " & REPT_month & "/" &  REPT_year & vbCr & "BOOLEAN - " & review_array(interview_const, incrementor_var)
' 					Else
' 						review_array(notes_const, incrementor_var) = "Unable to Access FS Review Information."
' 					End if
' 					Transmit 'to exit out of the pop-up screen
' 				End if
'
' 				If ma_case = True or msp_case = True then
' 					'read the HC review information
' 					Call write_value_and_transmit("X", 5, 71) 'HC Review Information
' 					EmReadscreen HC_review_popup, 20, 4, 32
' 					If HC_review_popup = "HEALTH CARE RENEWALS" then
' 					'The script will now read the CSR MO/YR and the Recert MO/YR
' 						EMReadScreen CSR_mo, 2, 8, 27   'IR dates
' 						EMReadScreen CSR_yr, 2, 8, 33
' 						If CSR_mo = "__" or CSR_yr = "__" then
' 							EMReadScreen CSR_mo, 2, 8, 71   'IR/AR dates
' 							EMReadScreen CSR_yr, 2, 8, 77
' 						End if
' 						EMReadScreen recert_mo, 2, 9, 27
' 						EMReadScreen recert_yr, 2, 9, 33
'
' 						HC_CSR_date = CSR_mo & "/" & CSR_yr
' 						If HC_CSR_date = "__/__" then HC_CSR_date = ""
'
' 						HC_ER_date = recert_mo & "/" & recert_yr
' 						If HC_ER_date = "__/__" then HC_ER_date = ""
'
' 						'Comparing CSR and ER daates to the month of REVS review
' 						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN
' 							review_array(current_SR_const, incrementor_var) = True
' 						ElseIf review_array(current_SR_const, incrementor_var) = "" Then
' 							review_array(current_SR_const, incrementor_var) = False
' 						End if
'
' 						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(no_interview_const, incrementor_var) = True
' 						If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) AND review_array(no_interview_const, incrementor_var) = "" then review_array(no_interview_const, incrementor_var) = False
'
' 						'Next HC ER and SR dates
' 						review_array(HC_next_SR_const, incrementor_var) = HC_CSR_date
' 						review_array(HC_next_ER_const, incrementor_var) = HC_ER_date
'
' 						Transmit 'to exit out of the pop-up screen
' 					Else
' 						Transmit 'to exit out of the pop-up screen
' 						review_array(notes_const, i) = "Unable to Access HC Review Information."
' 					End if
' 				End if
' 			End if
'
' 			'----------------------------------------------------------------------------------------------------language and Contact Information
' 			'Gathering the phone numbers
' 			call navigate_to_MAXIS_screen("STAT", "ADDR")
' 			EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
' 			If phone_number_one <> "( ___ ) ___ ____" then review_array(phone_1_const, incrementor_var) = phone_number_one
' 			EMReadScreen phone_number_two, 16, 18, 43
' 			If phone_number_two <> "( ___ ) ___ ____" then review_array(phone_2_const, incrementor_var) = phone_number_two
' 			EMReadScreen phone_number_three, 16, 19, 43
' 			If phone_number_three <> "( ___ ) ___ ____" then review_array(phone_3_const, incrementor_var) = phone_number_three
'
' 			'Going to STAT/MEMB for Language Information
' 			CALL navigate_to_MAXIS_screen("STAT", "MEMB")
' 			EMReadScreen interpreter_code, 1, 14, 68
' 			EMReadScreen language_coded, 16, 12, 46
' 			language_coded = replace(language_coded, "_", "")
' 			If trim(language_coded) = "" then
' 				EMReadScreen lang_ID, 2, 12, 42
' 				If lang_ID = "99" then lang_ID = "English"
' 				language_coded = lang_ID
' 			End if
'
' 			review_array(Interpreter_const, incrementor_var) = interpreter_code
' 			review_array(Language_const, incrementor_var) = language_coded
' 		End if
' 	End if
'
' 	Dialog1 = ""
' 	BeginDialog Dialog1, 0, 0, 286, 160, "Dialog"
' 	  ButtonGroup ButtonPressed
' 	    OkButton 230, 140, 50, 15
' 	  GroupBox 10, 10, 270, 50, "Programs"
' 	  Text 20, 25, 65, 10, "MFIP: " & review_array(MFIP_status_const, incrementor_var)
' 	  Text 20, 40, 65, 10, "DWP:" & review_array(DWP_status_const,  incrementor_var)
' 	  Text 115, 25, 65, 10, "GA:" & review_array(GA_status_const,   incrementor_var)
' 	  Text 115, 40, 65, 10, "MSA: " & review_array(MSA_status_const,  incrementor_var)
' 	  Text 205, 25, 65, 10, "GRH:" & review_array(GRH_status_const,  incrementor_var)
' 	  Text 205, 40, 65, 10, "SNAP: " & review_array(SNAP_status_const, incrementor_var)
' 	  GroupBox 10, 70, 270, 65, "REVW Detail"
' 	  ' If review_array(interview_const, incrementor_var) = True Then Text 20, 85, 200, 10, "Interview ER: TRUE"
' 	  ' If review_array(interview_const, incrementor_var) = False Then Text 20, 85, 200, 10, "Interview ER: FALSE"
' 	  ' If review_array(interview_const, incrementor_var) = "" Then Text 20, 85, 200, 10, "Interview ER: "
' 	  Text 20, 85, 200, 10, "Interview ER: " & review_array(interview_const, incrementor_var)
' 	  ' If review_array(no_interview_const, incrementor_var) = True Then Text 20, 100, 200, 10, "ER (no interview): TRUE"
' 	  ' If review_array(no_interview_const, incrementor_var) = False Then Text 20, 100, 200, 10, "ER (no interview): FALSE"
' 	  ' If review_array(no_interview_const, incrementor_var) = "" Then Text 20, 100, 200, 10, "ER (no interview):"
' 	  Text 20, 100, 200, 10, "ER (no interview):" & review_array(no_interview_const, incrementor_var)
' 	  ' If review_array(current_SR_const, incrementor_var) = True Then Text 20, 115, 200, 10, "CSR: TRUE"
' 	  ' If review_array(current_SR_const, incrementor_var) = False Then Text 20, 115, 200, 10, "CSR: FALSE"
' 	  ' If review_array(current_SR_const, incrementor_var) = "" Then Text 20, 115, 200, 10, "CSR: "
' 	  Text 20, 115, 200, 10, "CSR: " & review_array(current_SR_const, incrementor_var)
' 	EndDialog
'
' 	dialog Dialog1
'
' 	incrementor_var = incrementor_var + 1
' 	excel_row = ""
' Loop


'============================ OVERWRITE =================================='
'This is a script specific function and will not work outside of this script.
function read_case_details_for_review_report(incrementor_var)
	Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) 'function to check PRIV status
	If is_this_priv = True then
		review_array(notes_const, incrementor_var) = "PRIV Case."
		review_array(interview_const, incrementor_var) = ""
		review_array(no_interview_const, incrementor_var) = ""
		review_array(current_SR_const, incrementor_var) = ""
	Else
		EmReadscreen worker_prefix, 4, 21, 14
		If worker_prefix <> "X127" then
			review_array(notes_const, i) = "Out-of-County: " & right(worker_prefix, 2)
			review_array(notes_const, incrementor_var) = "PRIV Case."
			review_array(interview_const, incrementor_var) = ""
			review_array(no_interview_const, incrementor_var) = ""
			review_array(current_SR_const, incrementor_var) = ""
		Else
			'function to determine programs and the program's status---Yay Casey!
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
			
			If case_active = False then
				review_array(notes_const, incrementor_var) = "Case Not Active."
			Else
				'valuing the array variables from the inforamtion gathered in from CASE/CURR
				review_array(MFIP_status_const, incrementor_var) = mfip_case
				review_array(DWP_status_const,  incrementor_var) = dwp_case
				review_array(GA_status_const,   incrementor_var) = ga_case
				review_array(MSA_status_const,  incrementor_var) = msa_case
				review_array(GRH_status_const,  incrementor_var) = grh_case
				review_array(SNAP_status_const, incrementor_var) = snap_case
				review_array(MA_status_const,   incrementor_var) = ma_case
				review_array(MSP_status_const,  incrementor_var) = msp_case
				'----------------------------------------------------------------------------------------------------STAT/REVW
				CALL navigate_to_MAXIS_screen("STAT", "REVW")

				If family_cash_case = True or adult_cash_case = True or grh_case = True then
					'read the CASH review information
					Call write_value_and_transmit("X", 5, 35) 'CASH Review Information
					EmReadscreen cash_review_popup, 11, 5, 35
					If cash_review_popup = "GRH Reports" then
					'The script will now read the CSR MO/YR and the Recert MO/YR
						EMReadScreen CSR_mo, 2, 9, 26
						EMReadScreen CSR_yr, 2, 9, 32
						EMReadScreen recert_mo, 2, 9, 64
						EMReadScreen recert_yr, 2, 9, 70

						CASH_CSR_date = CSR_mo & "/" & CSR_yr
						If CASH_CSR_date = "__/__" then CASH_CSR_date = ""

						CASH_ER_date = recert_mo & "/" & recert_yr
						If CASH_ER_date = "__/__" then CASH_ER_date = ""

						'Comparing CSR dates to the month of REVS review
						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN review_array(current_SR_const, incrementor_var) = True

						'Determining if a case is ER, and if it meets interview requirement
						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then
							If mfip_case = True then review_array(interview_const, incrementor_var) = True             'MFIP interview requirement
							IF adult_cash_case = True or grh_case = True then review_array(no_interview_const, incrementor_var) = True    'Adult CASH programs do not meet interview requirement
						End if

						'Next CASH ER and SR dates
						review_array(CASH_next_SR_const, incrementor_var) = CASH_CSR_date
						review_array(CASH_next_ER_const, incrementor_var) = CASH_ER_date
					Else
						review_array(notes_const, incrementor_var) = "Unable to Access CASH Review Information."
					End if
					Transmit 'to exit out of the pop-up screen
				End if

				If snap_case = True then
					'read the SNAP review information
					Call write_value_and_transmit("X", 5, 58) 'SNAP Review Information
					EmReadscreen food_review_popup, 20, 5, 30
					If food_review_popup = "Food Support Reports" then
					'The script will now read the CSR MO/YR and the Recert MO/YR
						EMReadScreen CSR_mo, 2, 9, 26
						EMReadScreen CSR_yr, 2, 9, 32
						EMReadScreen recert_mo, 2, 9, 64
						EMReadScreen recert_yr, 2, 9, 70

						SNAP_CSR_date = CSR_mo & "/" & CSR_yr
						If SNAP_CSR_date = "__/__" then SNAP_CSR_date = ""

						SNAP_ER_date = recert_mo & "/" & recert_yr
						If SNAP_ER_date = "__/__" then SNAP_ER_date = ""

						'Comparing CSR and ER daates to the month of REVS review
						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN review_array(current_SR_const, incrementor_var) = True

						' If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) then review_array(interview_const, incrementor_var) = False
						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(interview_const, incrementor_var) = True

						'Next SNAP ER and SR dates
						review_array(SNAP_next_SR_const, incrementor_var) = SNAP_CSR_date
						review_array(SNAP_next_ER_const, incrementor_var) = SNAP_ER_date
					Else
						review_array(notes_const, incrementor_var) = "Unable to Access FS Review Information."
					End if
					Transmit 'to exit out of the pop-up screen
				End if

				If ma_case = True or msp_case = True then
					'read the HC review information
					Call write_value_and_transmit("X", 5, 71) 'HC Review Information
					EmReadscreen HC_review_popup, 20, 4, 32
					If HC_review_popup = "HEALTH CARE RENEWALS" then
					'The script will now read the CSR MO/YR and the Recert MO/YR
						EMReadScreen CSR_mo, 2, 8, 27   'IR dates
						EMReadScreen CSR_yr, 2, 8, 33
						If CSR_mo = "__" or CSR_yr = "__" then
							EMReadScreen CSR_mo, 2, 8, 71   'IR/AR dates
							EMReadScreen CSR_yr, 2, 8, 77
						End if
						EMReadScreen recert_mo, 2, 9, 27
						EMReadScreen recert_yr, 2, 9, 33

						HC_CSR_date = CSR_mo & "/" & CSR_yr
						If HC_CSR_date = "__/__" then HC_CSR_date = ""

						HC_ER_date = recert_mo & "/" & recert_yr
						If HC_ER_date = "__/__" then HC_ER_date = ""

						'Comparing CSR and ER daates to the month of REVS review
						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN review_array(current_SR_const, incrementor_var) = True

						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(no_interview_const, incrementor_var) = True

						'Next HC ER and SR dates
						review_array(HC_next_SR_const, incrementor_var) = HC_CSR_date
						review_array(HC_next_ER_const, incrementor_var) = HC_ER_date

						Transmit 'to exit out of the pop-up screen
					Else
						Transmit 'to exit out of the pop-up screen
						review_array(notes_const, i) = "Unable to Access HC Review Information."
					End if
				End if
			End if

			'----------------------------------------------------------------------------------------------------language and Contact Information
			'Gathering the phone numbers
			call navigate_to_MAXIS_screen("STAT", "ADDR")
			EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
			If phone_number_one <> "( ___ ) ___ ____" then review_array(phone_1_const, incrementor_var) = phone_number_one
			EMReadScreen phone_number_two, 16, 18, 43
			If phone_number_two <> "( ___ ) ___ ____" then review_array(phone_2_const, incrementor_var) = phone_number_two
			EMReadScreen phone_number_three, 16, 19, 43
			If phone_number_three <> "( ___ ) ___ ____" then review_array(phone_3_const, incrementor_var) = phone_number_three

			'Going to STAT/MEMB for Language Information
			CALL navigate_to_MAXIS_screen("STAT", "MEMB")
			EMReadScreen interpreter_code, 1, 14, 68
			EMReadScreen language_coded, 16, 12, 46
			language_coded = replace(language_coded, "_", "")
			If trim(language_coded) = "" then
				EMReadScreen lang_ID, 2, 12, 42
				If lang_ID = "99" then lang_ID = "English"
				language_coded = lang_ID
			End if

			review_array(Interpreter_const, incrementor_var) = interpreter_code
			review_array(Language_const, incrementor_var) = language_coded
		End if
	End if
end function


excel_row = 2
recert_cases = 0
update_count = 0
Do
	ReDim Preserve review_array(notes_const, recert_cases)	'This resizes the array based on if master notes were found or not
	review_array(worker_const,          recert_cases) = trim(ObjExcel.Cells(excel_row,  1).value)
	review_array(case_number_const,     recert_cases) = trim(ObjExcel.Cells(excel_row,  2).value)
	review_array(interview_const,       recert_cases) = trim(ObjExcel.Cells(excel_row,  3).value)      'COL C
	review_array(no_interview_const,    recert_cases) = trim(ObjExcel.Cells(excel_row,  4).value)      'COL D
	review_array(current_SR_const,      recert_cases) = trim(ObjExcel.Cells(excel_row,  5).value)      'COL E
	review_array(MFIP_status_const,     recert_cases) = ObjExcel.Cells(excel_row,  6).value      'COL F
	review_array(DWP_status_const,      recert_cases) = ObjExcel.Cells(excel_row,  7).value      'COL G
	review_array(GA_status_const,       recert_cases) = ObjExcel.Cells(excel_row,  8).value      'COL H
	review_array(MSA_status_const,      recert_cases) = ObjExcel.Cells(excel_row,  9).value      'COL I
	review_array(GRH_status_const,      recert_cases) = ObjExcel.Cells(excel_row, 10).value      'COL J
	review_array(CASH_next_SR_const,    recert_cases) = ObjExcel.Cells(excel_row, 11).value      'COL K
	review_array(CASH_next_ER_const,    recert_cases) = ObjExcel.Cells(excel_row, 12).value      'COL L
	review_array(SNAP_status_const,     recert_cases) = ObjExcel.Cells(excel_row, 13).value      'COL M
	review_array(SNAP_next_SR_const,    recert_cases) = ObjExcel.Cells(excel_row, 14).value      'COL N
	review_array(SNAP_next_ER_const,    recert_cases) = ObjExcel.Cells(excel_row, 15).value      'COL O
	review_array(MA_status_const,       recert_cases) = ObjExcel.Cells(excel_row, 16).value      'COL P
	review_array(MSP_status_const,      recert_cases) = ObjExcel.Cells(excel_row, 17).value      'COL Q
	review_array(HC_next_SR_const,      recert_cases) = ObjExcel.Cells(excel_row, 18).value      'COL R
	review_array(HC_next_ER_const,      recert_cases) = ObjExcel.Cells(excel_row, 19).value      'COL S
	review_array(Language_const,        recert_cases) = ObjExcel.Cells(excel_row, 20).value      'COL T
	review_array(Interpreter_const,     recert_cases) = ObjExcel.Cells(excel_row, 21).value      'COL U
	review_array(phone_1_const,         recert_cases) = ObjExcel.Cells(excel_row, 22).value      'COL V
	review_array(phone_2_const,         recert_cases) = ObjExcel.Cells(excel_row, 23).value      'COL W
	review_array(phone_3_const,         recert_cases) = ObjExcel.Cells(excel_row, 24).value      'COL X
	review_array(notes_const,           recert_cases) = ObjExcel.Cells(excel_row, 25).value      'COL Y

	MAXIS_case_number = review_array(case_number_const, recert_cases)
	' MsgBox "Excel Numb - " & excel_row & vbCr & "Case Number - " & MAXIS_case_number
	If MAXIS_case_number = "" Then Exit Do


	' If review_array(interview_const, recert_cases) = "False" Then MsgBox "INTV is FALSE"
	' If review_array(no_interview_const, recert_cases) = "False" Then MsgBox "NO intv is FALSE"
	' If review_array(current_SR_const, recert_cases) = "False" Then MsgBox "CSR is FALSE"

	If review_array(interview_const, recert_cases) = "False" AND review_array(no_interview_const, recert_cases) = "False" AND review_array(current_SR_const, recert_cases) = "False" Then
		' MsgBox "In it"
		' MsgBox "Excel Numb - " & excel_row & vbCr & "Case Number - " & MAXIS_case_number & vbCr & "INTV - " & review_array(interview_const, recert_cases) & vbCr & "No intv - " & review_array(no_interview_const, recert_cases) & vbCr & "CSR - " & review_array(current_SR_const, recert_cases)
		review_array(interview_const, recert_cases) = False
		review_array(no_interview_const, recert_cases) = False
		review_array(current_SR_const, recert_cases) = False

		Call read_case_details_for_review_report(recert_cases)

		ObjExcel.Cells(excel_row,  3).value = review_array(interview_const,       recert_cases)       'COL C
		ObjExcel.Cells(excel_row,  4).value = review_array(no_interview_const,    recert_cases)       'COL D
		ObjExcel.Cells(excel_row,  5).value = review_array(current_SR_const,      recert_cases)       'COL E

		ObjExcel.Range(ObjExcel.Cells(excel_row, 1), ObjExcel.Cells(excel_row, 31)).Interior.ColorIndex = 22

		update_count = update_count + 1
	End If

	excel_row = excel_row + 1
	recert_cases = recert_cases + 1
Loop
end_msg = "Done. Attempted to update " & update_count & " cases."
script_end_procedure(end_msg)
