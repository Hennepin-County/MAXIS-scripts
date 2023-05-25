'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - EX PARTE REPORT.vbs"
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
call changelog_update("05/10/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function find_unea_information()
	Call navigate_to_MAXIS_screen("STAT", "UNEA")
	For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
		EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = False
		MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = False

		EMReadScreen unea_vers, 1, 2, 78
		If unea_vers <> "0" Then
			Do
				EMReadScreen claim_num, 15, 6, 37
				EMReadScreen income_type_code, 2, 5, 37
				If income_type_code = "01" or income_type_code = "02" Then
					If left(start_of_claim, 9) <> MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) Then
						MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) = claim_num
					End If
				End if
				claim_num = replace(claim_num, "_", "")

				If income_type_code = "11" or income_type_code = "12" or income_type_code = "13" or income_type_code = "38" Then
					MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = True
					ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)

					VA_INCOME_ARRAY(va_case_numb_const, va_count) = MAXIS_case_number
					VA_INCOME_ARRAY(va_ref_numb_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					VA_INCOME_ARRAY(va_pers_name_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					VA_INCOME_ARRAY(va_pers_ssn_const, va_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					VA_INCOME_ARRAY(va_pers_pmi_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = income_type_code
					If income_type_code = "11" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Disability"
					If income_type_code = "12" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Pension"
					If income_type_code = "13" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Other"
					If income_type_code = "38" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Aid & Attendance"
					VA_INCOME_ARRAY(va_claim_numb_const, va_count) = claim_num
					EMReadScreen VA_INCOME_ARRAY(va_prosp_inc_const, va_count), 8, 18, 68
					VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = trim(VA_INCOME_ARRAY(va_prosp_inc_const, va_count))
					If VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "" Then VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "0.00"

					va_count = va_count + 1
				End If

				If income_type_code = "14" Then
					MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = True
					ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)

					UC_INCOME_ARRAY(uc_case_numb_const, uc_count) = MAXIS_case_number
					UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_pers_name_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = income_type_code
					UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
					UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) = claim_num
					EMReadScreen UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count), 8, 13, 68
					UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = trim(UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count))
					If UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "________" Then UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "0.00"

					uc_count = uc_count + 1
				End If

				transmit
				EMReadScreen next_unea_nav, 7, 24, 2
			Loop until next_unea_nav = "ENTER A"
		End If
	Next
end function

function get_list_of_members()
	client_count = UBound(MEMBER_INFO_ARRAY, 2) + 1
	If MEMBER_INFO_ARRAY(memb_pmi_numb_const, 0) = "" Then client_count = 0
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
	transmit
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadScreen client_PMI, 8, 4, 46
		client_PMI = trim(client_PMI)
		client_PMI = RIGHT("00000000" & client_PMI, 8)

		client_found = False
		For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
			If client_PMI = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then
				client_found = True
				EMReadScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 3, 4, 33
				EMReadScreen clt_age, 3, 8, 76
				MEMBER_INFO_ARRAY(memb_age_const, known_membs) = trim(clt_age)
				Exit For
			End If
		Next

		If client_found = False Then
			ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, client_count)

			EMReadScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, client_count), 3, 4, 33
			MEMBER_INFO_ARRAY(memb_pmi_numb_const, client_count) = client_PMI
			EMReadScreen SSN1, 3, 7, 42
			EMReadScreen SSN2, 2, 7, 46
			EMReadScreen SSN3, 4, 7, 49
			MEMBER_INFO_ARRAY(memb_ssn_const, client_count) = SSN1 & SSN2 & SSN3
			EMReadScreen clt_age, 3, 8, 76
			MEMBER_INFO_ARRAY(memb_age_const, client_count) = trim(clt_age)
			EMReadScreen last_name, 25, 6, 30
			EMReadScreen first_name, 12, 6, 63
			last_name = trim(replace(last_name, "_", ""))
			first_name = trim(replace(first_name, "_", ""))
			MEMBER_INFO_ARRAY(memb_name_const, client_count) = last_name & ", " & first_name
			MEMBER_INFO_ARRAY(memb_active_hc_const, client_count)	= False

			client_count = client_count + 1
		End If
		transmit
		EMReadScreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
end function

function send_sves_qury(ssn_or_claim, qury_finish)
	qury_finish = ""
	Call navigate_to_MAXIS_screen("INFC", "SVES")
	EMWriteScreen MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 68
	EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68
	Call write_value_and_transmit("QURY", 20, 70)										'Now we will enter the QURY screen to type the case number.

	If ssn_or_claim = "CLAIM" Then
		Call clear_line_of_text(5, 38)
		EMWriteScreen MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 7, 38
	End If
	EMWriteScreen MAXIS_case_number, 	11, 38
	EMWriteScreen "Y", 					14, 38
	transmit  'Now it sends the SVES.

	EMReadScreen duplicate_SVES, 	    7, 24, 2
	If duplicate_SVES = "WARNING" then transmit
	EMReadScreen confirm_SVES, 			6, 24, 2
	if confirm_SVES = "RECORD" then
		' PMI_array(SVES_status, item) = True
		qury_finish = date
	Else
		' PMI_array(SVES_status, item) = False
		qury_finish = "FAILED"
	END IF
end function

'END FUNCTIONS BLOCK =======================================================================================================


'DECLARATIONS ==============================================================================================================

Const memb_ref_numb_const 	= 0
Const memb_pmi_numb_const 	= 1
Const memb_ssn_const 		= 2
Const memb_age_const 		= 3
Const memb_name_const 		= 4
Const memb_active_hc_const	= 5
Const table_prog_1			= 6
Const table_type_1			= 7
Const table_prog_2			= 8
Const table_type_2			= 9
Const table_prog_3			= 10
Const table_type_3			= 11

Const unea_type_01_esists	= 20
Const unea_type_02_esists	= 21
Const unea_type_03_esists	= 22
Const unea_type_16_esists	= 23
Const unmatched_claim_numb	= 24
Const unea_VA_exists		= 25
Const unea_UC_exists		= 26

Const sves_qury_sent		= 40
Const second_qury_sent		= 41
Const sves_tpqy_response	= 42

Const memb_last_const 		= 70

Dim MEMBER_INFO_ARRAY()


Const va_case_numb_const 	= 0
Const va_ref_numb_const 	= 1
Const va_pers_name_const	= 2
Const va_pers_pmi_const		= 3
Const va_pers_ssn_const		= 4
Const va_inc_type_code_const 	= 5
Const va_inc_type_info_const	= 6
Const va_claim_numb_const 	= 7
Const va_prosp_inc_const 	= 8
Const va_last_const 		= 9

Dim VA_INCOME_ARRAY()
ReDim VA_INCOME_ARRAY(va_last_const, 0)

Const uc_case_numb_const 	= 0
Const uc_ref_numb_const 	= 1
Const uc_pers_name_const	= 2
Const uc_pers_pmi_const		= 3
Const uc_pers_ssn_const		= 4
Const uc_inc_type_code_const 	= 5
Const uc_inc_type_info_const	= 6
Const uc_claim_numb_const 	= 7
Const uc_prosp_inc_const 	= 8
Const uc_last_const 		= 9

Dim UC_INCOME_ARRAY()
ReDim UC_INCOME_ARRAY(va_last_const, 0)


'END DECLARATIONS BLOCK ====================================================================================================






'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

Confirm_Process_to_Run_btn	= 200
incorrect_process_btn		= 100
end_msg = "DONE"

If Day(date) < 1 Then ex_parte_function = "Prep"

'DISPLAYS DIALOG

DO
	DO
		DO
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 401, 255, "Ex Parte Report"
				DropListBox 300, 25, 90, 15, "Select one..."+chr(9)+"Prep"+chr(9)+"Phase 1"+chr(9)+"Phase 2", ex_parte_function
				ButtonGroup ButtonPressed
					OkButton 290, 235, 50, 15
					CancelButton 345, 235, 50, 15
				Text 5, 10, 400, 10, "This script will connect to the SQL Table to pull a list of cases to operate on based on the Ex Parte functionality selected."
				Text 200, 30, 95, 10, "Selection Ex Parte Function:"
				Text 10, 45, 35, 10, "Prep"
				Text 50, 45, 150, 10, "Timing - 4 Days before the BUDGET Month"
				Text 50, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
				Text 50, 65, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 50, 75, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
				Text 10, 90, 35, 10, "Phase 1"
				Text 50, 90, 135, 10, "Timing - 1st Day of the BUDGET Month"
				Text 50, 100, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 50, 110, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 50, 120, 125, 10, "Run each case through Background."
				Text 50, 130, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 50, 140, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
				Text 10, 155, 35, 10, "Phase 2"
				Text 50, 155, 160, 10, "Timing - 1st Day of the PROCESSING Month"
				Text 50, 165, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
				Text 50, 175, 145, 10, "Record in SQL Table any Updates found."
				Text 50, 185, 125, 10, "Run each case through Background."
				Text 50, 195, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 10, 215, 205, 10, "* * * * * THIS SCRIPT MUST BE RUN IN PRODUCTION * * * * *"
				Text 10, 235, 190, 10, "There is no CASE/NOTE entry by this script at this time."
			EndDialog

			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
			If ex_parte_function = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an Ex Parte Function."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP until err_msg = ""

		If ex_parte_function = "Prep" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)
			' ep_revw_mo = "07"
			' ep_revw_yr = "23"

		End If
		If ex_parte_function = "Phase 1" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 2, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 2, date)), 2)

		End If
		If ex_parte_function = "Phase 2" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 1, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 1, date)), 2)

		End If

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 341, 165, "Confirm Ex Parte process"
			EditBox 600, 700, 10, 10, fake_edit_box
			ButtonGroup ButtonPressed
				PushButton 10, 145, 210, 15, "CONFIRMED! This is the correct Process and Review Month", Confirm_Process_to_Run_btn
				PushButton 230, 145, 100, 15, "Incorrect Process/Month", incorrect_process_btn
			Text 10, 10, 225, 10, "You are running the Ex Parte Function " & ex_parte_function
			Text 10, 25, 190, 10, "This will run for the Ex Parte Review month of " & ep_revw_mo & "/" & ep_revw_yr
			If ex_parte_function = "Prep" Then
				GroupBox 5, 40, 240, 50, "Tasks to be Completed:"
				Text 20, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
				Text 20, 65, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 20, 75, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
			End If
			If ex_parte_function = "Phase 1" Then
				GroupBox 5, 40, 295, 70, "Tasks to be Completed:"
				Text 20, 55, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 20, 65, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 20, 75, 125, 10, "Run each case through Background."
				Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 20, 95, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
			End If
			If ex_parte_function = "Phase 2" Then
				GroupBox 5, 40, 305, 60, "Tasks to be Completed:"
				Text 20, 55, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
				Text 20, 65, 145, 10, "Record in SQL Table any Updates found."
				Text 20, 75, 125, 10, "Run each case through Background."
				Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
			End If
			Text 10, 115, 190, 10, "There is no CASE/NOTE entry by this script at this time."
			Text 10, 130, 330, 10, "Review the process datails and ex parte review month to confirm this is the correct run to complete."
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation

		If ButtonPressed = OK Then ButtonPressed = Confirm_Process_to_Run_btn

	Loop until ButtonPressed = Confirm_Process_to_Run_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in



If ex_parte_function = "Prep" Then
	' MsgBox  "At this point the script will pull the cases from a SQL Table that has identified cases due for a HC ER and evaluates them as potentially Ex Parte." & vbCr & vbCr &_
	' 		"If the case is potentially Ex Parte, the script will:" & vbCr &_
	' 		" - Send a SVES/QURY." & vbCr &_
	' 		" - Add the case to a report if VA Income is listed on the case to gather verification." & vbCr & vbCR &_
	' 		"This script will look at each case for the specified review month, preparing the case for review." & vbCr  & vbCr &_
	' 		"This script is run 4 business days before the Budget Month, or the end of the 3rd month BEFORE the ER month."

	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)

	va_count = 0
	uc_count = 0
	ex_parte_cases_count = 0

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table

	Do While NOT objRecordSet.Eof
		If IsNull(objRecordSet("PREP_Complete")) = True Then
			all_hc_is_ABD = ""
			SSA_income_exists = ""
			JOBS_income_exists = ""
			VA_income_exists = ""
			BUSI_income_exists = ""
			case_has_no_income = ""
			case_has_EPD = ""

			appears_ex_parte = True
			all_hc_is_ABD = True
			case_has_EPD = False
			case_is_in_henn = False
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
			' MsgBox MAXIS_case_number
			ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)
			memb_count = 0

			objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objELIGConnection = CreateObject("ADODB.Connection")
			Set objELIGRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objELIGRecordSet.Open objELIGSQL, objELIGConnection

			Do While NOT objELIGRecordSet.Eof

			' If objELIGRecordSet("MajorProgram") = NULL
				memb_known = False
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					If trim(objELIGRecordSet("PMINumber")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then
						memb_known = True
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= objELIGRecordSet("EligType")
						ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= objELIGRecordSet("EligType")
						ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= objELIGRecordSet("EligType")
						End If
						If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
						If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
						If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
						If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
						If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
						If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
						If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

						If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False
						If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True
					End If
				Next

				If memb_known = False Then
					ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)

					MEMBER_INFO_ARRAY(memb_pmi_numb_const, memb_count) 	= trim(objELIGRecordSet("PMINumber"))
					MEMBER_INFO_ARRAY(memb_ssn_const, memb_count) 		= trim(objELIGRecordSet("SocialSecurityNbr"))
					MEMBER_INFO_ARRAY(memb_name_const, memb_count) 		= trim(objELIGRecordSet("Name"))
					MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
					MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(objELIGRecordSet("MajorProgram"))
					MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(objELIGRecordSet("EligType"))

					If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
					If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
					If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
					If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
					If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
					If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
					If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

					If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False
					If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True

					memb_count = memb_count + 1
				End if
				objELIGRecordSet.MoveNext
			Loop
			objELIGRecordSet.Close
			objELIGConnection.Close
			Set objELIGRecordSet=nothing
			Set objELIGConnection=nothing


			SSA_income_exists = False
			RR_income_exists = False
			VA_income_exists = False
			UC_income_exists = False
			PRISM_income_exists = False
			Other_UNEA_income_exists = False
			JOBS_income_exists = False
			BUSI_income_exists = False

			objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objIncomeConnection = CreateObject("ADODB.Connection")
			Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

			Do While NOT objIncomeRecordSet.Eof
				If objIncomeRecordSet("IncExpTypeCode") = "UNEA" Then
					If objIncomeRecordSet("IncomeTypeCode") = "01" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "02" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "03" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "16" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "11" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "12" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "13" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "38" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "14" Then UC_income_exists = True

					If objIncomeRecordSet("IncomeTypeCode") = "36" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "37" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "39" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "40" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "36" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "37" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "39" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "40" Then Other_UNEA_income_exists = True

					If objIncomeRecordSet("IncomeTypeCode") = "06" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "15" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "17" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "18" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "23" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "24" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "25" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "26" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "27" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "28" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "29" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "08" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "35" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "43" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "47" Then Other_UNEA_income_exists = True
				End If
				If objIncomeRecordSet("IncExpTypeCode") = "JOBS" Then JOBS_income_exists = True
				If objIncomeRecordSet("IncExpTypeCode") = "BUSI" Then BUSI_income_exists = True

				objIncomeRecordSet.MoveNext
			Loop
			objIncomeRecordSet.Close
			objIncomeConnection.Close
			Set objIncomeRecordSet=nothing
			Set objIncomeConnection=nothing


			If appears_ex_parte = True Then

				'check HC ER date in STAT/REVW
				Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
				If is_this_priv = True Then appears_ex_parte = False
				If is_this_priv = False Then
					Call write_value_and_transmit("X", 5, 71)
					EMReadScreen STAT_HC_ER_mo, 2, 8, 27
					EMReadScreen STAT_HC_ER_yr, 2, 8, 33
					If ep_revw_mo <> STAT_HC_ER_mo or ep_revw_yr <> STAT_HC_ER_yr Then  appears_ex_parte = False
				End If
			End If

			If appears_ex_parte = True Then
				Call navigate_to_MAXIS_screen("STAT", "MEMB")
				Call get_list_of_members

				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
				EMReadScreen case_pw, 7, 21, 14
				If left(case_pw, 4) = "X127" Then case_is_in_henn = True
				If case_is_in_henn = False then  appears_ex_parte = False
				If case_active = False Then appears_ex_parte = False
				If ma_status <> "ACTIVE" and msp_status <> "ACTIVE" Then appears_ex_parte = False
				' If msp_status <> "ACTIVE" Then appears_ex_parte = False

				' If mfip_status = "ACTIVE" OR snap_status = "ACTIVE" Then
			End If

			case_has_no_income = False
			If SSA_income_exists = False and RR_income_exists = False and VA_income_exists = False and UC_income_exists = False and PRISM_income_exists = False and Other_UNEA_income_exists = False and JOBS_income_exists = False and BUSI_income_exists = False Then case_has_no_income = True

			If appears_ex_parte = True Then
				If Other_UNEA_income_exists = True OR JOBS_income_exists = True OR BUSI_income_exists = True Then
					If mfip_status = "ACTIVE" Then
						' If all_hc_is_ABD = True and case_has_EPD = False Then
						appears_ex_parte = True
					ElseIf snap_status = "ACTIVE" Then
						'find income
						appears_ex_parte = True
					Else
						' If Other_UNEA_income_exists = True Then appears_ex_parte = False
						' If JOBS_income_exists = True Then appears_ex_parte = False
						' If BUSI_income_exists = True Then appears_ex_parte = False
						appears_ex_parte = False
					End If
				End If
			End If

			If appears_ex_parte = True Then ex_parte_cases_count = ex_parte_cases_count + 1

			' test_string = ""
			' For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
			' 	test_string = test_string & vbcr & "MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const,known_membs) & " - " & MEMBER_INFO_ARRAY(memb_name_const, known_membs) & " - ONE: " & MEMBER_INFO_ARRAY(table_prog_1, known_membs) & "-" & MEMBER_INFO_ARRAY(table_type_1, known_membs) & " - TWO: " & MEMBER_INFO_ARRAY(table_prog_2, known_membs) & "-" & MEMBER_INFO_ARRAY(table_type_2, known_membs) & " - THREE: " & MEMBER_INFO_ARRAY(table_prog_3, known_membs) & "-" & MEMBER_INFO_ARRAY(table_type_3, known_membs)
			' Next
			' test_string = test_string & vbcr & "STAT Review Month: " & STAT_HC_ER_mo & "/" & STAT_HC_ER_yr
			' If MEMBER_INFO_ARRAY(memb_name_const, 0) = "" Then
			' 	Call navigate_to_MAXIS_screen("CASE", "PERS")
			' 	MsgBox "CASE/PERS" & vbCr & MAXIS_case_number & vbCr & test_string & vbCr & vbCr & "appears_ex_parte - " & appears_ex_parte
			' 	Call navigate_to_MAXIS_screen("STAT", "REVW")
			' 	MsgBox "STAT/REVW" & vbCr & MAXIS_case_number & vbCr & test_string & vbCr & vbCr & "appears_ex_parte - " & appears_ex_parte
			' End If
			' If UBound(MEMBER_INFO_ARRAY, 2) > 0 Then
			' 	Call navigate_to_MAXIS_screen("CASE", "PERS")
			' 	MsgBox "Multiple Members" & vbCr & MAXIS_case_number & vbCr & test_string & vbCr & vbCr & "appears_ex_parte - " & appears_ex_parte
			' End If

			' 'Read from ELIG List and Income list to determine if Ex Parte and update the table
			' 'Create a list of the people on HC in the case in an array with income information
			' 'ELIG Type is EX, DX, BX - this is possible
			' 'If income included is something other than SSI, , RSDI, RR, or VA

			' If appears_ex_parte = True Then
			' 	'check the case to make sure HC is active for the case AND for the persons
			' 	Call write_value_and_transmit("PERS", 20, 69)

			' 	pers_row = 10
			' 	hc_stat = ""
			' 	Do
			' 		EMReadScreen read_pmi, 8, pers_row, 34
			' 		read_pmi = trim(read_pmi)

			' 		If read_pmi = PERS_PMI Then						'NEED to deal with multiple people
			' 			EMReadScreen ref_number, 2, pers_row, 3
			' 			EMReadScreen hc_stat, 1, pers_row, 61
			' 			Exit Do
			' 		End If
			' 		pers_row = pers_row + 3
			' 		If pers_row = 19 Then
			' 			PF8
			' 			pers_row = 10
			' 			EMReadScreen end_of_list, 9, 24, 14
			' 			If end_of_list = "LAST PAGE" Then Exit Do
			' 		End If
			' 	Loop until read_pmi = ""



			' End If

			' 'If NOT Ex Parte - update SQL with False for Ex Parte and Code PREP column as Not Ex Parte
			' If appears_ex_parte = False Then objRecordSet.Open "UPDATE"

			If appears_ex_parte = True Then
				'For each case that is indicated as potentially ExParte, we are going to take preperation actions
				last_va_count = va_count
				last_uc_count = uc_count

				'Find if there is a claim number that is not associated with the persons SSN
				'Read if VA income is on UNEA to add that person to a list to verify VA
				'Read if UC income is on UNEA to add that person to a list to verify UC
				Call find_unea_information

				Call back_to_SELF

				'Send a SVES/CURY for all persons on a case
				Call navigate_to_MAXIS_screen("INFC", "SVES")
				'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
				EMReadScreen agreement_check, 9, 2, 24
				IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

				'We need to loop through each HH Member on the case and send a QURY for every one.
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					Call send_sves_qury("SSN", qury_finish)
					MEMBER_INFO_ARRAY(sves_qury_sent, each_memb) = qury_finish

					objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) & "%'"

					'Creating objects for Access
					Set objIncomeConnection = CreateObject("ADODB.Connection")
					Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					'This is the file path for the statistics Access database.
					' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
					objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					MEMBER_INFO_ARRAY(second_qury_sent, each_memb) = False
					If MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) <> "" Then
						Call send_sves_qury("CLAIM", qury_finish)
						MEMBER_INFO_ARRAY(second_qury_sent, each_memb) = qury_finish

						objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) & "%'"

						'Creating objects for Access
						Set objIncomeConnection = CreateObject("ADODB.Connection")
						Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

						'This is the file path for the statistics Access database.
						' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
						objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					End If
				Next

				'If SSA income is listed on the INCOME LIST Table, we update the item in the table when the QURY goes out

				If va_count = 0 Then
					If last_va_count <> va_count Then
						If last_va_count = 0 Then
							va_excel_created = True
							'Create an Excel file to record members that have VA Income
							Set objVAExcel = CreateObject("Excel.Application")
							objVAExcel.Visible = True
							Set objWorkbook = objVAExcel.Workbooks.Add()
							objVAExcel.DisplayAlerts = True

							'Setting the first 4 col as worker, case number, name, and APPL date
							objVAExcel.Cells(1, 1).Value = "CASE NUMBER"
							objVAExcel.Cells(1, 2).Value = "REF"
							objVAExcel.Cells(1, 3).Value = "NAME"
							objVAExcel.Cells(1, 4).Value = "PMI NUMBER"
							objVAExcel.Cells(1, 5).Value = "SSN"
							objVAExcel.Cells(1, 6).Value = "VA INC TYPE"
							objVAExcel.Cells(1, 7).Value = "VA CLAIM NUMB"
							objVAExcel.Cells(1, 8).Value = "CURR VA INCOME"
							objVAExcel.Cells(1, 9).Value = "Verified VA Income"

							FOR i = 1 to 9		'formatting the cells'
								objVAExcel.Cells(1, i).Font.Bold = True		'bold font'
							NEXT

							va_excel_row = 2
							va_inc_count = 0
						End If

						Do
							objVAExcel.Cells(va_excel_row, 1).value = VA_INCOME_ARRAY(va_case_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 2).value = VA_INCOME_ARRAY(va_ref_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 3).value = VA_INCOME_ARRAY(va_pers_name_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 4).value = VA_INCOME_ARRAY(va_pers_pmi_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 5).value = VA_INCOME_ARRAY(va_pers_ssn_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 6).value = VA_INCOME_ARRAY(va_claim_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 7).value = VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) & " - " & VA_INCOME_ARRAY(va_inc_type_info_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 8).value = VA_INCOME_ARRAY(va_prosp_inc_const, va_inc_count)

							va_inc_count = va_inc_count + 1
							va_excel_row = va_excel_row + 1
						Loop until va_inc_count = va_count
					End If
				End If

				If uc_count <> 0 Then
					If last_uc_count <> uc_count Then
						If last_uc_count = 0 Then
							uc_excel_created = True
							'Create an Excel file to record members that have VA Income
							Set objUCExcel = CreateObject("Excel.Application")
							objUCExcel.Visible = True
							Set objWorkbook = objUCExcel.Workbooks.Add()
							objUCExcel.DisplayAlerts = True

							'Setting the first 4 col as worker, case number, name, and APPL date
							objUCExcel.Cells(1, 1).Value = "CASE NUMBER"
							objUCExcel.Cells(1, 2).Value = "REF"
							objUCExcel.Cells(1, 3).Value = "NAME"
							objUCExcel.Cells(1, 4).Value = "PMI NUMBER"
							objUCExcel.Cells(1, 5).Value = "SSN"
							objUCExcel.Cells(1, 6).Value = "UC INC TYPE"
							objUCExcel.Cells(1, 7).Value = "UC CLAIM NUMB"
							objUCExcel.Cells(1, 8).Value = "CURR UC INCOME"
							objUCExcel.Cells(1, 9).Value = "Verified UC Income"

							FOR i = 1 to 9		'formatting the cells'
								objUCExcel.Cells(1, i).Font.Bold = True		'bold font'
							NEXT

							uc_excel_row = 2
							uc_inc_count = 0
						End If

						Do
							objUCExcel.Cells(uc_excel_row, 1).value = UC_INCOME_ARRAY(uc_case_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 2).value = UC_INCOME_ARRAY(uc_ref_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 3).value = UC_INCOME_ARRAY(uc_pers_name_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 4).value = UC_INCOME_ARRAY(uc_pers_pmi_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 5).value = UC_INCOME_ARRAY(uc_pers_ssn_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 6).value = UC_INCOME_ARRAY(uc_claim_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 7).value = UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) & " - " & UC_INCOME_ARRAY(uc_inc_type_info_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 8).value = UC_INCOME_ARRAY(uc_prosp_inc_const, uc_inc_count)

							uc_inc_count = uc_inc_count + 1
							uc_excel_row = uc_excel_row + 1
						Loop until uc_inc_count = uc_count
					End If
				End If

				'save details of the actions into the table
				' objRecordSet.Open "UPDATE"
			End If

			Call back_to_SELF
			' objRecordSet.Update "SelectExParte", appears_ex_parte
			prep_status = date
			If appears_ex_parte = False Then
				prep_status = "Not Ex Parte"
				' If mfip_status = "ACTIVE" OR snap_status = "ACTIVE" Then prep_status = "SNAP/MFIP"
			End If

			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', PREP_Complete = '" & prep_status & "', AllHCisABD = '" & all_hc_is_ABD & "', SSAIncomExist = '" & SSA_income_exists & "', WagesExist = '" & JOBS_income_exists & "', VAIncomeExist = '" & VA_income_exists & "', SelfEmpExists = '" & BUSI_income_exists & "', NoIncome = '" & case_has_no_income & "', EPDonCase = '" & case_has_EPD & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objUpdateConnection = CreateObject("ADODB.Connection")
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			' objUpdateRecordSet.Close
			' objUpdateConnection.Close
			' Set objUpdateRecordSet=nothing
			' Set objUpdateConnection=nothing

		End If
		objRecordSet.MoveNext
	Loop

	For col_to_autofit = 1 to 9
		If va_excel_created = True Then objVAExcel.columns(col_to_autofit).AutoFit()
		If uc_excel_created = True Then objUCExcel.columns(col_to_autofit).AutoFit()
	Next
	' MsgBox "Cases that appear Ex Parte: " & ex_parte_cases_count

    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing


	end_msg = "BULK Prep Run has been completed."

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
	percent_ex_parte = ex_parte_count/case_count
	percent_ex_parte = percent_ex_parte * 100
	percent_ex_parte = FormatNumber(percent_ex_parte, 2, -1, 0, -1)

	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count
	end_msg = end_msg & vbCr & "This appears to be " & percent_ex_parte & "% of cases."
End If

If ex_parte_function = "Phase 1" Then
	MsgBox 	"In preparation for the HSR completion of a Phase 1 review, the script will complete updates to MAXIS information, to prevent HSRs from having to amnually enter verified information." & vbCr & vbCr &_
			"If the case is potentially Ex Parte, the script will:" & vbCr &_
			" - Read SVES/TPQY" & vbCr &_
			" - Update UNEA and MEDI with SSA information from SVES/TPQY." & vbCr &_
			" - Enter VA Income reported back after verification." & vbCr &_
			" - Create a CASE/NOTE of any information verified and updated in MAXIS." & vbCr &_
			" - Run the case through background." & vbCr &_
			" - Capture details of the income verified and the Eligibility results into the Table to track Ex Parte work." & vbCr & vbCr &_
			"This script will look at each case for the specified review month, preparing the case to be assigned to an HSR for Phase 1 Review of Ex Parte Eligbility." & vbCr  & vbCr &_
			"This script is run on the 1st of the month of the Budget Month."
	'Open the CASE LIST Table
	'Loop through each item on the CASE LIST Table
		'For each case that is indicated as Ex parte, we are going to update the case information

		'Read SVES/TPQY for all persons on a case
		'Update MAXIS UNEA panels with information from TPQY
		'Update MAXIS MEDI panels with information from TPQY

		'Update MAXIS UNEA panels with information from the VA Verifications report

		'CASE/NOTE details of the case information

		'Send the case through background
		'Read ELIG and MMIS

		'Save all details from the income updates and ELIG information into the SQL Table
End If

If ex_parte_function = "Phase 2" Then
	MsgBox "Phase 2 BULK Run Details to be added later. This functionality will prep cases for HSR Review at Phase 2, which will happen at the beginning of the Processing month (the month before the Review Month)."
End If


'Loop through all the SQL Items and look for the right revew month and year and phase to determine if it's done.

Call script_end_procedure(end_msg)
