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

function find_last_approved_ELIG_version(cmd_row, cmd_col, version_number, version_date, version_result)
	Call write_value_and_transmit("99", cmd_row, cmd_col)

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

	Call write_value_and_transmit(elig_version, 18, 54)
	version_number = "0" & elig_version
	version_date = elig_date
	version_result = elig_result
end function

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Connects to BlueZone
EMConnect ""

'Checking for MAXIS
Call check_for_MAXIS(True)


file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\Cash pending from 6-1.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)

excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)

	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

	ObjExcel.Cells(excel_row, 8).Value = unknown_cash_pending
	ObjExcel.Cells(excel_row, 9).Value = ga_status
	ObjExcel.Cells(excel_row, 10).Value = msa_status
	ObjExcel.Cells(excel_row, 11).Value = mfip_status
	ObjExcel.Cells(excel_row, 12).Value = dwp_status

	Call back_to_SELF
	excel_row = excel_row + 1
Loop until trim(ObjExcel.Cells(excel_row, 2).Value) = ""

script_end_procedure("Thanks! We're done here.")

file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases June 3.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)

excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)

	call navigate_to_MAXIS_screen("ELIG", "GRH ")
	Call write_value_and_transmit("GRFB", 20, 71)
	' EMWriteScreen "06", 20, 55
	' EMWriteScreen "22", 20, 58
	' ' transmit
	' Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result)

	' EMReadScreen grh_elig_memb_elig_type_code, 2, 6, 53
	' EMReadScreen grh_elig_case_test_fail_file, 				6, 11, 8
	' EMReadScreen grh_elig_case_test_assets, 				6, 8, 45
	' EMReadScreen grh_elig_case_test_verif, 					6, 13, 45
	' EMReadScreen grh_elig_budg_vendor_number, 	8, 6, 25

	' If grh_elig_memb_elig_type_code = "01" Then  grh_elig_memb_elig_type_info = "SSI"
	' If grh_elig_memb_elig_type_code = "02" Then  grh_elig_memb_elig_type_info = "MFIP"
	' If grh_elig_memb_elig_type_code = "03" Then  grh_elig_memb_elig_type_info = "Blind"
	' If grh_elig_memb_elig_type_code = "04" Then  grh_elig_memb_elig_type_info = "Disabled"
	' If grh_elig_memb_elig_type_code = "05" Then  grh_elig_memb_elig_type_info = "Aged"
	' If grh_elig_memb_elig_type_code = "06" Then  grh_elig_memb_elig_type_info = "Adult"
	' If grh_elig_memb_elig_type_code = "07" Then  grh_elig_memb_elig_type_info = "None"
	' If grh_elig_memb_elig_type_code = "08" Then  grh_elig_memb_elig_type_info = "Residential Treatment"
	' MsgBox grh_elig_memb_elig_type_code & " - " & grh_elig_memb_elig_type_info
	EMReadScreen grh_elig_budg_vendor_number_one, 	8, 6, 25
	EMReadScreen grh_elig_budg_vendor_number_two, 	8, 6, 44
	EMReadScreen grh_elig_budg_vendor_number_thr, 	25, 6, 54

	ObjExcel.Cells(excel_row, 19).Value = trim(grh_elig_budg_vendor_number_one)
	ObjExcel.Cells(excel_row, 20).Value = trim(grh_elig_budg_vendor_number_two)
	ObjExcel.Cells(excel_row, 21).Value = trim(grh_elig_budg_vendor_number_thr)

	Call back_to_SELF
	excel_row = excel_row + 1
Loop until trim(ObjExcel.Cells(excel_row, 2).Value) = ""


script_end_procedure("Thanks! We're done here.")
