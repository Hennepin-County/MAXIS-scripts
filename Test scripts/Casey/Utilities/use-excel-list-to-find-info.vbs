'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "BULK - REPT-ELIG LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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


'This function is used to grab all active X numbers according to the supervisor X number(s) inputted
FUNCTION create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
	'Getting to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")
	'Sorting by supervisor
	PF5
	PF5
	'Reseting array_name
	array_name = ""
	'Splitting the list of inputted supervisors...
	supervisor_array = replace(supervisor_array, " ", "")
	supervisor_array = split(supervisor_array, ",")
	FOR EACH unit_supervisor IN supervisor_array
		IF unit_supervisor <> "" THEN
			'Entering the supervisor number and sending a transmit
			CALL write_value_and_transmit(unit_supervisor, 21, 12)
			MAXIS_row = 7
			DO
				EMReadScreen worker_ID, 8, MAXIS_row, 5
				worker_ID = trim(worker_ID)
				IF worker_ID = "" THEN EXIT DO
				array_name = trim(array_name & " " & worker_ID)
				MAXIS_row = MAXIS_row + 1
				IF MAXIS_row = 19 THEN
					PF8
					EMReadScreen end_check, 9, 24,14
					If end_check = "LAST PAGE" Then Exit Do
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
END FUNCTION

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result, approval_found)
	Call write_value_and_transmit("99", cmd_row, cmd_col)
	approval_found = True

	row = 7
	Do
		EMReadScreen elig_version, 2, row, 22
		EMReadScreen elig_date, 8, row, 26
		EMReadScreen elig_result, 10, row, 37
		EMReadScreen approval_status, 10, row, 50

		elig_version = trim(elig_version)
		elig_result = trim(elig_result)
		approval_status = trim(approval_status)

		If approval_status = "APPROVED" Then Exit Do

		row = row + 1
	Loop until approval_status = ""

	Call clear_line_of_text(18, 54)
	If approval_status = "" Then
		approval_found = false
		PF3
	Else
		Call write_value_and_transmit(elig_version, 18, 54)
		version_number = "0" & elig_version
		version_date = elig_date
		version_result = elig_result
	End If
end function

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Connects to BlueZone
EMConnect ""

'Checking for MAXIS
Call check_for_MAXIS(True)


' file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\Cash pending from 6-1.xlsx"
' visible_status = True
' alerts_status = True
' Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
'
' excel_row = 2
' Do
' 	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)
'
' 	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
'
' 	ObjExcel.Cells(excel_row, 8).Value = unknown_cash_pending
' 	ObjExcel.Cells(excel_row, 9).Value = ga_status
' 	ObjExcel.Cells(excel_row, 10).Value = msa_status
' 	ObjExcel.Cells(excel_row, 11).Value = mfip_status
' 	ObjExcel.Cells(excel_row, 12).Value = dwp_status
'
' 	Call back_to_SELF
' 	excel_row = excel_row + 1
' Loop until trim(ObjExcel.Cells(excel_row, 2).Value) = ""
'
' script_end_procedure("Thanks! We're done here.")

file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases June 3.xlsx"
file_url = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Auto-Close Pause Project\Tier Two Review\Tier Two Completed Review Data.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Expedited Determination\Exp Exch Approval Review.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Expedited Determination\Cases Still Pending from Exp Exch.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases Aug 9.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\Cash Cases 8-1-22.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases Sept 16.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
' Call navigate_to_MAXIS_screen("CCOL", "CLIC")
MAXIS_footer_month = "09"
MAXIS_footer_year = "22"
excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)

	call navigate_to_MAXIS_screen("ELIG", "GRH ")
	EmReadScreen at_grh_elig, 16, 2, 33
	If at_grh_elig = "GRH ELIG Results" Then
		' EMWriteScreen MAXIS_footer_month, 20, 55
		' EMWriteScreen MAXIS_footer_year, 20, 58
		' transmit
		Call find_last_approved_ELIG_version(20, 79, version_number, version_date, version_result, approval_found)

		EMReadScreen grh_elig_type, 2, 6, 53
		EMReadScreen grh_elig_case_test_assets, 6, 8, 45
		EMReadScreen grh_elig_case_test_fail_file, 6, 11, 8
		EMReadScreen grh_elig_case_test_verif, 6, 13, 45
		EMReadScreen grh_elig_case_test_income, 6, 11, 45
		' MsgBox "grh_elig_type - " & grh_elig_type
		' MsgBox "grh_elig_case_test_assets - " & grh_elig_case_test_assets & vbCr & "grh_elig_case_test_fail_file - " & grh_elig_case_test_fail_file & vbCr & "grh_elig_case_test_verif - " & grh_elig_case_test_verif


		' If grh_elig_type = "01" Then  grh_elig_type_info = "SSI"
	 	' If grh_elig_type = "02" Then  grh_elig_type_info = "MFIP"
	 	' If grh_elig_type = "03" Then  grh_elig_type_info = "Blind"
	 	' If grh_elig_type = "04" Then  grh_elig_type_info = "Disabled"
	 	' If grh_elig_type = "05" Then  grh_elig_type_info = "Aged"
	 	' If grh_elig_type = "06" Then  grh_elig_type_info = "Adult"
	 	' If grh_elig_type = "07" Then  grh_elig_type_info = "None"
	 	' If grh_elig_type = "08" Then  grh_elig_type_info = "Residential Treatment"
		'
		' Call write_value_and_transmit("GRFB", 20, 71)
		'
	 	' EMReadScreen grh_elig_budg_vendor_number_one, 	8, 6, 25
	 	' EMReadScreen grh_elig_budg_vendor_number_two, 	8, 6, 44
		' MsgBox "grh_elig_budg_vendor_number_one - " & grh_elig_budg_vendor_number_one
		ObjExcel.Cells(excel_row, 5).Value = grh_elig_case_test_assets
		ObjExcel.Cells(excel_row, 6).Value = grh_elig_case_test_fail_file
		ObjExcel.Cells(excel_row, 7).Value = grh_elig_case_test_verif
		ObjExcel.Cells(excel_row, 8).Value = grh_elig_case_test_income
	End If

	' deny_cash_dwp_reason_info = ""
	' deny_cash_mfip_reason_info = ""
	' deny_cash_msa_reason_info = ""
	' deny_cash_ga_reason_info = ""
	' MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)
	' MAXIS_case_number = left(MAXIS_case_number & "       ", 8)
	' Call write_value_and_transmit(MAXIS_case_number, 4, 8)
	'
	' call navigate_to_MAXIS_screen("ELIG", "DENY")
	' EMWriteScreen "08", 19, 54
	' EMWriteScreen "22", 19, 57
	' Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
	' transmit
	'
	' EMReadScreen deny_cash_dwp_reason_code, 2, 8, 46
	' EMReadScreen deny_cash_mfip_reason_code, 2, 9, 46
	' EMReadScreen deny_cash_msa_reason_code, 2, 12, 46
	' EMReadScreen deny_cash_ga_reason_code, 2, 13, 46
	'
	' If deny_cash_dwp_reason_code = "" Then deny_cash_dwp_reason_info = ""
	' If deny_cash_dwp_reason_code = "01" Then deny_cash_dwp_reason_info = "No Eligible Child"
	' If deny_cash_dwp_reason_code = "02" Then deny_cash_dwp_reason_info = "Application Withdrawn"
	' If deny_cash_dwp_reason_code = "03" Then deny_cash_dwp_reason_info = "Initial Income"
	' If deny_cash_dwp_reason_code = "04" Then deny_cash_dwp_reason_info = "Assets"
	' If deny_cash_dwp_reason_code = "05" Then deny_cash_dwp_reason_info = "Fail To Cooperate"
	' If deny_cash_dwp_reason_code = "06" Then deny_cash_dwp_reason_info = "Child Support Disqualification"
	' If deny_cash_dwp_reason_code = "07" Then deny_cash_dwp_reason_info = "Employment Services Disqualification"
	' If deny_cash_dwp_reason_code = "08" Then deny_cash_dwp_reason_info = "Death"
	' If deny_cash_dwp_reason_code = "09" Then deny_cash_dwp_reason_info = "Residence"
	' If deny_cash_dwp_reason_code = "10" Then deny_cash_dwp_reason_info = "Transfer of Resources"
	' If deny_cash_dwp_reason_code = "11" Then deny_cash_dwp_reason_info = "Verification"
	' If deny_cash_dwp_reason_code = "12" Then deny_cash_dwp_reason_info = "Strike"
	' If deny_cash_dwp_reason_code = "13" Then deny_cash_dwp_reason_info = "Program Active"
	' If deny_cash_dwp_reason_code = "14" Then deny_cash_dwp_reason_info = "4 Month Limit"
	' If deny_cash_dwp_reason_code = "15" Then deny_cash_dwp_reason_info = "MFIP Conversion"
	' If deny_cash_dwp_reason_code = "23" Then deny_cash_dwp_reason_info = "Duplicate Assistance"
	' If deny_cash_dwp_reason_code = "99" Then deny_cash_dwp_reason_info = "PND2 Denial"
	' If deny_cash_dwp_reason_code = "TL" Then deny_cash_dwp_reason_info = "TANF Time Limit"
	'
	' If deny_cash_mfip_reason_code = "" Then deny_cash_mfip_reason_info = ""
	' If deny_cash_mfip_reason_code = "01" Then deny_cash_mfip_reason_info = "No Eligible Child"
	' If deny_cash_mfip_reason_code = "02" Then deny_cash_mfip_reason_info = "Application Withdrawn"
	' If deny_cash_mfip_reason_code = "03" Then deny_cash_mfip_reason_info = "Initial Income"
	' If deny_cash_mfip_reason_code = "04" Then deny_cash_mfip_reason_info = "Monthly Income"
	' If deny_cash_mfip_reason_code = "05" Then deny_cash_mfip_reason_info = "Assets"
	' If deny_cash_mfip_reason_code = "06" Then deny_cash_mfip_reason_info = "Fail To Cooperate"
	' If deny_cash_mfip_reason_code = "07" Then deny_cash_mfip_reason_info = "Fail To Cooperate with IEVS"
	' If deny_cash_mfip_reason_code = "08" Then deny_cash_mfip_reason_info = "Death"
	' If deny_cash_mfip_reason_code = "09" Then deny_cash_mfip_reason_info = "Residence"
	' If deny_cash_mfip_reason_code = "10" Then deny_cash_mfip_reason_info = "Transfer of Resources"
	' If deny_cash_mfip_reason_code = "11" Then deny_cash_mfip_reason_info = "Verification"
	' If deny_cash_mfip_reason_code = "12" Then deny_cash_mfip_reason_info = "Strike"
	' If deny_cash_mfip_reason_code = "13" Then deny_cash_mfip_reason_info = "Fail To File"
	' If deny_cash_mfip_reason_code = "14" Then deny_cash_mfip_reason_info = "Program Active"
	' If deny_cash_mfip_reason_code = "23" Then deny_cash_mfip_reason_info = "Duplicate Assistance"
	' If deny_cash_mfip_reason_code = "24" Then deny_cash_mfip_reason_info = "Minor Living Arrangement"
	' If deny_cash_mfip_reason_code = "TL" Then deny_cash_mfip_reason_info = "TANF Time Limit"
	' If deny_cash_mfip_reason_code = "33" Then deny_cash_mfip_reason_info = "Diversionary Work Program"
	' If deny_cash_mfip_reason_code = "34" Then deny_cash_mfip_reason_info = "Sanction Period"
	' If deny_cash_mfip_reason_code = "35" Then deny_cash_mfip_reason_info = "Sanction Date Compliance"
	' If deny_cash_mfip_reason_code = "99" Then deny_cash_mfip_reason_info = "PND2 Denial System Entered"
	'
	' If deny_cash_msa_reason_code = "" Then deny_cash_msa_reason_info = ""
	' If deny_cash_msa_reason_code = "01" Then deny_cash_msa_reason_info = "No Eligible Member"
	' If deny_cash_msa_reason_code = "03" Then deny_cash_msa_reason_info = "Verification"
	' If deny_cash_msa_reason_code = "08" Then deny_cash_msa_reason_info = "Application Withdrawn"
	' If deny_cash_msa_reason_code = "10" Then deny_cash_msa_reason_info = "Residence"
	' If deny_cash_msa_reason_code = "11" Then deny_cash_msa_reason_info = "Assets"
	' If deny_cash_msa_reason_code = "24" Then deny_cash_msa_reason_info = "Program Active"
	' If deny_cash_msa_reason_code = "28" Then deny_cash_msa_reason_info = "Fail To File"
	' If deny_cash_msa_reason_code = "29" Then deny_cash_msa_reason_info = "Applicant Eligible"
	' If deny_cash_msa_reason_code = "30" Then deny_cash_msa_reason_info = "Prospective Gross Income"
	' If deny_cash_msa_reason_code = "31" Then deny_cash_msa_reason_info = "Prospective Net Income"
	' If deny_cash_msa_reason_code = "99" Then deny_cash_msa_reason_info = "PND2 Denial System Entered"
	'
	' If deny_cash_ga_reason_code = "" Then deny_cash_ga_reason_info = ""
	' If deny_cash_ga_reason_code = "01" Then deny_cash_ga_reason_info = "No Eligible Person"
	' If deny_cash_ga_reason_code = "02" Then deny_cash_ga_reason_info = "Net Income"
	' If deny_cash_ga_reason_code = "03" Then deny_cash_ga_reason_info = "Verification"
	' If deny_cash_ga_reason_code = "04" Then deny_cash_ga_reason_info = "Non Cooperation"
	' If deny_cash_ga_reason_code = "06" Then deny_cash_ga_reason_info = "Other Benefits"
	' If deny_cash_ga_reason_code = "07" Then deny_cash_ga_reason_info = "Address Unknown"
	' If deny_cash_ga_reason_code = "08" Then deny_cash_ga_reason_info = "Application Withdrawn"
	' If deny_cash_ga_reason_code = "09" Then deny_cash_ga_reason_info = "Client Request"
	' If deny_cash_ga_reason_code = "10" Then deny_cash_ga_reason_info = "Residence"
	' If deny_cash_ga_reason_code = "11" Then deny_cash_ga_reason_info = "Assets"
	' If deny_cash_ga_reason_code = "12" Then deny_cash_ga_reason_info = "Transfer of Resource"
	' If deny_cash_ga_reason_code = "14" Then deny_cash_ga_reason_info = "Interim Assistance Agreement"
	' If deny_cash_ga_reason_code = "15" Then deny_cash_ga_reason_info = "Out Of County"
	' If deny_cash_ga_reason_code = "16" Then deny_cash_ga_reason_info = "Disqualify"
	' If deny_cash_ga_reason_code = "17" Then deny_cash_ga_reason_info = "Interview"
	' If deny_cash_ga_reason_code = "19" Then deny_cash_ga_reason_info = "Fail to File"
	' If deny_cash_ga_reason_code = "21" Then deny_cash_ga_reason_info = "Duplicate Assistance"
	' If deny_cash_ga_reason_code = "22" Then deny_cash_ga_reason_info = "Death"
	' If deny_cash_ga_reason_code = "23" Then deny_cash_ga_reason_info = "Eligible Other Benefits"
	' If deny_cash_ga_reason_code = "26" Then deny_cash_ga_reason_info = "Program Active"
	' If deny_cash_ga_reason_code = "29" Then deny_cash_ga_reason_info = "Lump Sum"
	' If deny_cash_ga_reason_code = "99" Then deny_cash_ga_reason_info = "PND2 Denial System Entered"
	'
	' ObjExcel.Cells(excel_row, 8).Value = deny_cash_dwp_reason_info
	' ObjExcel.Cells(excel_row, 9).Value = deny_cash_mfip_reason_info
	' ObjExcel.Cells(excel_row, 10).Value = deny_cash_msa_reason_info
	' ObjExcel.Cells(excel_row, 11).Value = deny_cash_ga_reason_info

	' Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
	'
	' ObjExcel.Cells(excel_row, 8).Value = unknown_cash_pending
	' ObjExcel.Cells(excel_row, 9).Value = mfip_status
	' ObjExcel.Cells(excel_row, 10).Value = dwp_status
	' ObjExcel.Cells(excel_row, 11).Value = ga_status
	' ObjExcel.Cells(excel_row, 12).Value = msa_status
	' ObjExcel.Cells(excel_row, 13).Value = grh_status

	' ObjExcel.Cells(excel_row, 4).Value = snap_status
	Call Back_to_SELF
	' EMWriteScreen "08", 20, 42
	' EMWriteScreen "22", 20, 46
	'
	'
	' Call navigate_to_MAXIS_screen("STAT", "PDED")
	' EMReadScreen shelter_special_need, 1, 18, 78
	' If shelter_special_need = "Y" Then
	' 	ObjExcel.Cells(excel_row, 19).Value = "TRUE"
	'
	' 	Call navigate_to_MAXIS_screen("STAT", "HEST")
	' 	EMReadScreen heat_ac_amount, 6, 13, 75
	' 	EMReadScreen elec_amount, 6, 14, 75
	' 	EMReadScreen phone_amount, 6, 15, 75
	' 	heat_ac_amount = trim(heat_ac_amount)
	' 	elec_amount = trim(elec_amount)
	' 	phone_amount = trim(phone_amount)
	' 	If heat_ac_amount = "" Then heat_ac_amount = 0
	' 	If elec_amount = "" Then elec_amount = 0
	' 	If phone_amount = "" Then phone_amount = 0
	' 	heat_ac_amount = heat_ac_amount*1
	' 	elec_amount = elec_amount*1
	' 	phone_amount = phone_amount*1
	' 	hest_total = heat_ac_amount + elec_amount + phone_amount
	' 	ObjExcel.Cells(excel_row, 22).Value = hest_total
	'
	' 	Call navigate_to_MAXIS_screen("STAT", "SHEL")
	' 	EMreadScreen shel_rent, 8, 11, 56
	' 	EMreadScreen shel_lot_rent, 8, 12, 56
	' 	EMreadScreen shel_mortgage, 8, 13, 56
	' 	EMreadScreen shel_insur, 8, 14, 56
	' 	EMreadScreen shel_taxes, 8, 15, 56
	' 	EMreadScreen shel_room, 8, 16, 56
	' 	EMreadScreen shel_garage, 8, 17, 56
	'
	' 	shel_rent = trim(shel_rent)
	' 	shel_lot_rent = trim(shel_lot_rent)
	' 	shel_mortgage = trim(shel_mortgage)
	' 	shel_insur = trim(shel_insur)
	' 	shel_taxes = trim(shel_taxes)
	' 	shel_room = trim(shel_room)
	' 	shel_garage = trim(shel_garage)
	'
	' 	shel_rent = replace(shel_rent, "_", "")
	' 	shel_lot_rent = replace(shel_lot_rent, "_", "")
	' 	shel_mortgage = replace(shel_mortgage, "_", "")
	' 	shel_insur = replace(shel_insur, "_", "")
	' 	shel_taxes = replace(shel_taxes, "_", "")
	' 	shel_room = replace(shel_room, "_", "")
	' 	shel_garage = replace(shel_garage, "_", "")
	'
	' 	If shel_rent = "" Then shel_rent = 0
	' 	If shel_lot_rent = "" Then shel_lot_rent = 0
	' 	If shel_mortgage = "" Then shel_mortgage = 0
	' 	If shel_insur = "" Then shel_insur = 0
	' 	If shel_taxes = "" Then shel_taxes = 0
	' 	If shel_room = "" Then shel_room = 0
	' 	If shel_garage = "" Then shel_garage = 0
	'
	' 	shel_rent = shel_rent*1
	' 	shel_lot_rent = shel_lot_rent*1
	' 	shel_mortgage = shel_mortgage*1
	' 	shel_insur = shel_insur*1
	' 	shel_taxes = shel_taxes*1
	' 	shel_room = shel_room*1
	' 	shel_garage = shel_garage*1
	'
	' 	shel_total = shel_rent + shel_lot_rent + shel_mortgage + shel_insur + shel_taxes + shel_room + shel_garage
	' 	ObjExcel.Cells(excel_row, 21).Value = shel_total
	' End If
	'
	'
	' call navigate_to_MAXIS_screen("ELIG", "MSA ")
	' EMWriteScreen "08", 20, 56
	' EMWriteScreen "22", 20, 59
	' Call find_last_approved_ELIG_version(20, 79, version_number, version_date, version_result, approved_version_found)
	'
	' transmit
	' transmit
	' EMReadScreen msa_elig_case_budg_type, 12, 3, 25
	' EMReadScreen msa_elig_net_income, 9, 7, 72
	' Call write_value_and_transmit("X", 6, 43)
	' EMReadScreen msa_elig_budg_special_needs, 10, 17, 59
	' transmit
	' PF3
	' PF3
	'
	' ObjExcel.Cells(excel_row, 17).Value = version_date
	' ObjExcel.Cells(excel_row, 18).Value = msa_elig_budg_special_needs
	' ObjExcel.Cells(excel_row, 20).Value = trim(msa_elig_net_income)
	' ObjExcel.Cells(excel_row, 23).Value = trim(msa_elig_case_budg_type)
	'
	' call navigate_to_MAXIS_screen("ELIG", "MSA ")
	' EMWriteScreen "09", 20, 56
	' EMWriteScreen "22", 20, 59
	' Call find_last_approved_ELIG_version(20, 79, version_number, version_date, version_result, approved_version_found)
	'
	' ObjExcel.Cells(excel_row, 16).Value = version_date

	' Call back_to_SELF
	'
	'
	' ObjExcel.Cells(excel_row, 104).Value = "NO"
	' month_of_op = #2/1/2022#
	'
	' prog_col = 105
	' number_col = 106
	' balance_col = 107
	'
	' ccol_row = 8
	' Do
	' 	EMReadScreen claim_pd_start, 5, ccol_row, 26
	' 	EMReadScreen claim_pd_end, 5, ccol_row, 32
	' 	If trim(claim_pd_start) <> "" Then
	' 		claim_pd_start = replace(claim_pd_start, "/", "/1/")
	' 		claim_pd_end = replace(claim_pd_end, "/", "/1/")
	' 		claim_pd_start = DateAdd("d", 0, claim_pd_start)
	' 		claim_pd_end = DateAdd("d", 0, claim_pd_end)
	'
	'
	' 		If DateDiff("d", month_of_op, claim_pd_start) <= 0 AND DateDiff("d", month_of_op, claim_pd_end) >= 0 Then
	'
	'
	'
	' 		' If claim_pd = "02/22" Then
	' 			' MsgBox "in the time range" & vbCr & vbCr &"claim_pd_start - " & claim_pd_start & vbCr & "date diff - " & DateDiff("d", month_of_op, claim_pd_start) & vbCr & "claim_pd_end - " & claim_pd_end & vbCr & "date diff - " & DateDiff("d", month_of_op, claim_pd_end)
	' 			EMReadScreen claim_prog, 2, ccol_row, 5
	' 			EMReadScreen claim_numb, 6, ccol_row, 54
	' 			EMReadScreen claim_bal, 9, ccol_row, 38
	'
	' 			ObjExcel.Cells(excel_row, 104).Value = "YES"
	'
	' 			ObjExcel.Cells(excel_row, prog_col).Value = trim(claim_prog)
	' 			ObjExcel.Cells(excel_row, number_col).Value = trim(claim_numb)
	' 			ObjExcel.Cells(excel_row, balance_col).Value = trim(claim_bal)
	'
	' 			prog_col = prog_col + 3
	' 			number_col = number_col + 3
	' 			balance_col = balance_col + 3
	'
	' 		End If
	' 	End If
	'
	'
	' 	ccol_row = ccol_row + 1
	' 	If ccol_row = 18 Then
	' 		exit do
	' 		ObjExcel.Cells(excel_row, 109).Value = "???"
	' 	End If
	' 	EMReadScreen next_claim_pd, 5, ccol_row, 26
	' Loop until trim(next_claim_pd) = ""
	excel_row = excel_row + 1
Loop until trim(ObjExcel.Cells(excel_row, 1).Value) = ""

MsgBox "STOP HERE"

Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)
	MAXIS_case_number = left(MAXIS_case_number & "       ", 8)
	Call write_value_and_transmit(MAXIS_case_number, 20, 39)
	' call navigate_to_MAXIS_screen("ELIG", "HC  ")
	elig_hc_row = 8
	excel_col = 19
	Do
		EMReadScreen hc_prog, 4, elig_hc_row, 28
		hc_prog = trim(hc_prog)
		If hc_prog = "MA" Then
			Call write_value_and_transmit("X", elig_hc_row, 26)
			EMReadScreen first_month, 5, 6, 19
			first_month = replace(first_month, "/", "/1/")
			PF3
			ObjExcel.Cells(excel_row, excel_col).Value = first_month
			excel_col = excel_col + 1
		End If
		' If hc_prog <> "" AND hc_prog <> "NO V" AND hc_prog <> "NO R" Then
		' 	ObjExcel.Cells(excel_row, excel_col).Value = hc_prog
		' 	excel_col = excel_col + 1
		' End If
		elig_hc_row = elig_hc_row + 1
	Loop until hc_prog = ""

	' call navigate_to_MAXIS_screen("ELIG", "GRH ")
	' Call write_value_and_transmit("GRFB", 20, 71)
	' ' EMWriteScreen "06", 20, 55
	' ' EMWriteScreen "22", 20, 58
	' ' ' transmit
	' ' Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result)
	'
	' ' EMReadScreen grh_elig_memb_elig_type_code, 2, 6, 53
	' ' EMReadScreen grh_elig_case_test_fail_file, 				6, 11, 8
	' ' EMReadScreen grh_elig_case_test_assets, 				6, 8, 45
	' ' EMReadScreen grh_elig_case_test_verif, 					6, 13, 45
	' ' EMReadScreen grh_elig_budg_vendor_number, 	8, 6, 25
	'
	' ' If grh_elig_memb_elig_type_code = "01" Then  grh_elig_memb_elig_type_info = "SSI"
	' ' If grh_elig_memb_elig_type_code = "02" Then  grh_elig_memb_elig_type_info = "MFIP"
	' ' If grh_elig_memb_elig_type_code = "03" Then  grh_elig_memb_elig_type_info = "Blind"
	' ' If grh_elig_memb_elig_type_code = "04" Then  grh_elig_memb_elig_type_info = "Disabled"
	' ' If grh_elig_memb_elig_type_code = "05" Then  grh_elig_memb_elig_type_info = "Aged"
	' ' If grh_elig_memb_elig_type_code = "06" Then  grh_elig_memb_elig_type_info = "Adult"
	' ' If grh_elig_memb_elig_type_code = "07" Then  grh_elig_memb_elig_type_info = "None"
	' ' If grh_elig_memb_elig_type_code = "08" Then  grh_elig_memb_elig_type_info = "Residential Treatment"
	' ' MsgBox grh_elig_memb_elig_type_code & " - " & grh_elig_memb_elig_type_info
	' EMReadScreen grh_elig_budg_vendor_number_one, 	8, 6, 25
	' EMReadScreen grh_elig_budg_vendor_number_two, 	8, 6, 44
	' EMReadScreen grh_elig_budg_vendor_number_thr, 	25, 6, 54
	'
	' ObjExcel.Cells(excel_row, 19).Value = trim(grh_elig_budg_vendor_number_one)
	' ObjExcel.Cells(excel_row, 20).Value = trim(grh_elig_budg_vendor_number_two)
	' ObjExcel.Cells(excel_row, 21).Value = trim(grh_elig_budg_vendor_number_thr)

	' Call back_to_SELF
	excel_row = excel_row + 1
Loop until trim(ObjExcel.Cells(excel_row, 2).Value) = ""


script_end_procedure("Thanks! We're done here.")
