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
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases Oct 6.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases Nov 8.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\All Cases Dec 6.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
' Call navigate_to_MAXIS_screen("CCOL", "CLIC")
' count = 1
' xl_col = 5
' Do
' 	ObjExcel.Cells(1, xl_col).Value = "ELIGIBILITY - " & count
' 	ObjExcel.Cells(1, xl_col+1).Value = "ELIG TYPE - " & count
' 	ObjExcel.Cells(1, xl_col+2).Value = "INCOME - " & count
' 	ObjExcel.Cells(1, xl_col+3).Value = "PERSON and PROG - " & count
' 	ObjExcel.Cells(1, xl_col+4).Value = "APPROVED? - " & count
'
'
' 	ObjExcel.Columns(xl_col).AutoFit()
' 	ObjExcel.Columns(xl_col+1).AutoFit()
' 	ObjExcel.Columns(xl_col+2).AutoFit()
' 	ObjExcel.Columns(xl_col+3).AutoFit()
'  	ObjExcel.Columns(xl_col+4).AutoFit()
'
' 	xl_col = xl_col + 5
' 	count = count + 1
' Loop Until count = 8


MAXIS_footer_month = "12"
MAXIS_footer_year = "22"
excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)

	Call navigate_to_MAXIS_screen("CASE", "CURR")
	EMreadScreen current_pw, 7, 21, 14
	EMreadScreen current_co, 2, 21, 16
	ObjExcel.Cells(excel_row, 4).Value = current_pw
	ObjExcel.Cells(excel_row, 5).Value = current_co

	' xl_col = 5
	'
	' call navigate_to_MAXIS_screen("ELIG", "HC  ")
	' EMWriteScreen MAXIS_footer_month, 19, 54
	' EMWriteScreen MAXIS_footer_year, 19, 57
	' transmit
	'
	' hc_row = 8
	' Do
	' 	EMReadScreen new_hc_elig_ref_numbs, 2, hc_row, 3
	' 	EMReadScreen new_hc_elig_full_name, 17, hc_row, 7
	'
	' 	If new_hc_elig_ref_numbs = "  " Then
	' 		new_hc_elig_ref_numbs = hc_elig_ref_numbs
	' 		new_hc_elig_full_name = hc_elig_full_name
	' 	End If
	' 	hc_elig_ref_numbs = new_hc_elig_ref_numbs
	' 	hc_elig_full_name = new_hc_elig_full_name
	'
	' 	hc_elig_full_name = trim(hc_elig_full_name)
	'
	' 	EMReadScreen clt_hc_prog, 4, hc_row, 28
	' 	If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "    " Then
	'
	' 		EMReadScreen prog_status, 3, hc_row, 68
	' 		If prog_status <> "APP" Then                        'Finding the approved version
	' 			EMReadScreen total_versions, 2, hc_row, 64
	' 			If total_versions = "01" Then
	' 				hc_prog_elig_appd = False
	' 			Else
	' 				EMReadScreen current_version, 2, hc_row, 58
	' 				' MsgBox "hc_row - " & hc_row & vbCr & "current_version - " & current_version
	' 				If current_version = "01" Then
	' 					hc_prog_elig_appd = False
	' 				Else
	' 					prev_version = right ("00" & abs(current_version) - 1, 2)
	' 					EMWriteScreen prev_version, hc_row, 58
	' 					transmit
	' 					hc_prog_elig_appd = True
	' 				End If
	'
	' 			End If
	' 		Else
	' 			hc_prog_elig_appd = True
	' 		End If
	' 	Else
