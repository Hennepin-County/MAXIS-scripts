'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - BULK REVW ER List for Waived Interviews.vbs"
start_time = timer

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

'array
'case number
'cash er t/f
'snap er t/f
'cash prog active
'possible waiver interview t/f

case_nbr_col = 2
run_another_script(t_drive & "\Eligibility Support\Scripts\Script Files\reviews-delayed.vbs")
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

const case_nbr_const 			= 0
const cash_er_tf_const 			= 1
const snap_er_tf_const 			= 2
const cash_prog_const 			= 3
const caf_date_const 			= 4
const intv_date_const 			= 5
const poss_waived_intv_tf_const = 6
const excel_row_const 			= 7
const case_notes_const 			= 8
' const _const =

Dim ER_CASES_ARRAY()
ReDim ER_CASES_ARRAY(case_notes_const, 0)

EMConnect ""
Call check_for_MAXIS(true)

'Identify which month it is for revw
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 181, 55, "Which Month"
  DropListBox 95, 15, 80, 45, "Select One..."+chr(9)+"October"+chr(9)+"November"+chr(9)+"December"+chr(9)+"January"+chr(9)+"February", month_to_check
  ButtonGroup ButtonPressed
    OkButton 125, 35, 50, 15
  Text 10, 20, 80, 10, "What is the ER month?"
EndDialog

dialog Dialog1

'Read the list of all ERS for the month - pull cases into an array
excel_file_path = ""
If month_to_check = "October" Then
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\2020\10-20 Renewals.xlsx"
	REVS_footer_mo = "10"
	REVS_footer_yr = "20"
	sheet_to_select = "ER cases 10-20"
	' sheet_to_select = "TRIAL"
	ADJUSTED_CASES_ARRAY = oct_revw_to_adjust_array
End If
If month_to_check = "November" Then
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\2020\11-20 Renewals.xlsx"
	REVS_footer_mo = "11"
	REVS_footer_yr = "20"
	sheet_to_select = "ER cases 11-20"
	' sheet_to_select = "TRIAL"
	ADJUSTED_CASES_ARRAY = nov_revw_to_adjust_array
End If
If month_to_check = "December" Then
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\2020\12-20 Renewals.xlsx"
	REVS_footer_mo = "12"
	REVS_footer_yr = "20"
	sheet_to_select = "ER cases 12-20"
	' sheet_to_select = "TRIAL"
	ADJUSTED_CASES_ARRAY = dec_revw_to_adjust_array
End If
If month_to_check = "January" Then
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\2021\01-21 Renewals.xlsx"
	REVS_footer_mo = "01"
	REVS_footer_yr = "21"
	sheet_to_select = "ER cases 01-21"
	' sheet_to_select = "TRIAL"
	ADJUSTED_CASES_ARRAY = jan_revw_to_adjust_array
End If
If month_to_check = "February" Then
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\2021\02-21 Renewals.xlsx"
	REVS_footer_mo = "02"
	REVS_footer_yr = "21"
	sheet_to_select = "ER cases 02-21"
	' sheet_to_select = "TRIAL"
	ADJUSTED_CASES_ARRAY = feb_revw_to_adjust_array
End If

'Initial Dialog which requests a file path for the excel file
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 361, 70, "Select the File Path"
  EditBox 130, 25, 175, 15, excel_file_path
  ButtonGroup ButtonPressed
    PushButton 310, 25, 45, 15, "Browse...", select_a_file_button
    OkButton 305, 50, 50, 15
  Text 10, 10, 170, 10, "Select the File Path:"
  Text 10, 30, 120, 10, "Select an Excel file for recert cases:"
EndDialog


'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
	Dialog Dialog1
	If ButtonPressed = cancel then stopscript
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(excel_file_path, ".xlsx")
Loop until ButtonPressed = OK and select_a_file_button <> ""

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

'Activates worksheet based on user selection
ObjExcel.worksheets(sheet_to_select).Activate

excel_row = 2
case_entry = 0
'reading each line of the Excel file and adding case number information to the array
Do
	If month_to_check = "October" Then
		If ObjExcel.Cells(excel_row, 32).Value = "" AND ObjExcel.Cells(excel_row, 33).Value <> "" Then
			ReDim Preserve ER_CASES_ARRAY(case_notes_const, case_entry)
			ER_CASES_ARRAY(case_nbr_const, case_entry) = trim(objExcel.Cells(excel_row, case_nbr_col).Value)
			ER_CASES_ARRAY(excel_row_const, case_entry) = excel_row
			ER_CASES_ARRAY(poss_waived_intv_tf_const, case_entry) = TRUE

			case_entry = case_entry + 1
		End If
	Else
	    ReDim Preserve ER_CASES_ARRAY(case_notes_const, case_entry)
	    ER_CASES_ARRAY(case_nbr_const, case_entry) = trim(objExcel.Cells(excel_row, case_nbr_col).Value)
	    ER_CASES_ARRAY(excel_row_const, case_entry) = excel_row
		ER_CASES_ARRAY(poss_waived_intv_tf_const, case_entry) = TRUE

	    case_entry = case_entry + 1
	End If
    excel_row = excel_row + 1
    next_case_number = trim(objExcel.Cells(excel_row, case_nbr_col).Value)
loop until next_case_number = ""
last_excel_row = excel_row -1

total_cases = case_entry

'Compare to the adjusted array - if on adjested array - possible waiver is false
test = 0
For each adjusted_revw_case in ADJUSTED_CASES_ARRAY
	For er_case = 0 To UBound(ER_CASES_ARRAY, 2)
		If ER_CASES_ARRAY(case_nbr_const, er_case) = adjusted_revw_case Then
			ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
			Exit For
		End If
	Next
	test = test + 1
Next
MsgBox "Checked adjusted review cases - " & test

Set objRange = objExcel.Range("I1").EntireColumn
objRange.Insert(xlShiftToRight)                             'inserting one column to the end of the data in the spreadsheet
objRange.Insert(xlShiftToRight)                             'inserting one column to the end of the data in the spreadsheet
ObjExcel.Cells(1, 9).Value = "Poss Waived Intv"
ObjExcel.Cells(1, 10).Value = "Cash Prog Actv"

' If month_to_check = "October" Then
' 	MAXIS_footer_month = "10"
' 	MAXIS_footer_year = "20"
' 	call back_to_SELF
' 	For er_case = 0 to UBound(ER_CASES_ARRAY, 2)
' 		MAXIS_case_number = ER_CASES_ARRAY(case_nbr_const, er_case)
' 		If ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE Then
' 			Call navigate_to_MAXIS_screen("STAT", "REVW")
'
' 			EMReadScreen cash_grh_er_code, 1, 7, 40
' 			EMReadScreen snap_er_code, 1, 7, 60
' 			EMReadScreen revs_caf_date, 8, 13, 37
' 			EMReadScreen revs_intv_date, 8, 15, 37
'
' 			If cash_grh_er_code = "-" OR cash_grh_er_code = "_" OR cash_grh_er_code = " " Then
' 				ER_CASES_ARRAY(cash_er_tf_const, er_case) = FALSE
' 			Else
' 				ER_CASES_ARRAY(cash_er_tf_const, er_case) = TRUE
' 			End If
' 			If snap_er_code = "-" OR snap_er_code = "_" OR snap_er_code = " " Then
' 				ER_CASES_ARRAY(snap_er_tf_const, er_case) = FALSE
' 				ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
' 			Else
' 				ER_CASES_ARRAY(snap_er_tf_const, er_case) = TRUE
' 				ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE
' 				' MsgBox "Case Number - " & ER_CASES_ARRAY(case_nbr_const, er_case) & vbCr & "Cash REVW - " & cash_grh_er_code & vbCr & "SNAP REVW - " & snap_er_code & vbCr & "CAF DATE - " & revs_caf_date & vbCr & "INTERVIEW DATE - " & revs_intv_date & vbCr & ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case)
' 			End If
' 			If revs_caf_date <> "__ __ __" Then ER_CASES_ARRAY(caf_date_const, er_case) = replace(revs_caf_date, " ", "/")
' 			If revs_intv_date <> "__ __ __" Then
' 				ER_CASES_ARRAY(intv_date_const, er_case) = replace(revs_intv_date, " ", "/")
' 				ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
' 			End If
' 			' MsgBox "Case Number - " & ER_CASES_ARRAY(case_nbr_const, er_case) & vbCr & "Cash REVW - " & cash_grh_er_code & vbCr & "SNAP REVW - " & snap_er_code & vbCr & "CAF DATE - " & revs_caf_date & vbCr & "INTERVIEW DATE - " & revs_intv_date & vbCr & ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case)
'
' 		End If
'
' 		If ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE Then
' 			If ER_CASES_ARRAY(cash_er_tf_const, er_case) = TRUE Then
'
' 				Call navigate_to_MAXIS_screen("STAT", "PROG")
' 				cash_one_prog = ""
' 				cash_two_prog = ""
'
' 				EMReadScreen cash_one_prog_status, 4, 6, 74
' 				EMReadScreen cash_two_prog_status, 4, 7, 74
' 				EMReadScreen grh_prog_status, 4, 9, 74
'
' 				If cash_one_prog_status = "ACTV" Then EMReadScreen cash_one_prog, 2, 6, 67
' 				If cash_two_prog_status = "ACTV" Then EMReadScreen cash_two_prog, 2, 6, 67
'
' 				If cash_one_prog = "MF" THen cash_one_prog = "MFIP"
' 				If cash_one_prog = "MS" THen cash_one_prog = "MSA"
'
' 				If cash_two_prog = "MF" THen cash_two_prog = "MFIP"
' 				If cash_two_prog = "MS" THen cash_two_prog = "MSA"
'
' 				If cash_one_prog = "MFIP" OR cash_two_prog = "MFIP" Then ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
'
' 				If cash_one_prog <> "" Then ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", " & cash_one_prog
' 				If cash_two_prog <> "" Then ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", " & cash_two_prog
' 				If grh_prog_status = "ACTV" THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", GRH"
' 				If left(ER_CASES_ARRAY(cash_prog_const, er_case), 2) = ", " Then ER_CASES_ARRAY(cash_prog_const, er_case) = right(ER_CASES_ARRAY(cash_prog_const, er_case), len(ER_CASES_ARRAY(cash_prog_const, er_case)) - 2)
' 			End If
' 		End If
' 		Call Back_to_self
' 		row = ER_CASES_ARRAY(excel_row_const, er_case)
' 		ObjExcel.Cells(row, 9).Value = ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) & ""
' 		ObjExcel.Cells(row, 10).Value = ER_CASES_ARRAY(cash_prog_const, er_case)
' 	Next
' 	script_end_procedure("OCT done now")
' End If
'
' 'ONLY possible waiver is NOT false then read on
' Call Back_to_self
' Call navigate_to_MAXIS_screen("REPT", "REVS")
' EMWriteScreen REVS_footer_mo, 20, 55
' EMWriteScreen REVS_footer_yr, 20, 58
' transmit
'
' For er_case = 0 to UBound(ER_CASES_ARRAY, 2)
' 	If ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE Then
' 		EMWriteScreen ER_CASES_ARRAY(case_nbr_const, er_case), 20, 33
' 		transmit
' 		' MsgBox "Pause - " & ER_CASES_ARRAY(case_nbr_const, er_case)
'
'
' 		row = 1
' 		col = 1
' 		EMSearch ER_CASES_ARRAY(case_nbr_const, er_case), row, col
'
' 		If row < 20 and row > 0 Then
' 			EMReadScreen cash_grh_er_code, 1, row, 39
' 			EMReadScreen snap_er_code, 1, row, 45
' 			EMReadScreen revs_caf_date, 8, row, 62
' 			EMReadScreen revs_intv_date, 8, row, 72
'
' 			If cash_grh_er_code = "-" OR cash_grh_er_code = "_" OR cash_grh_er_code = " " Then
' 				ER_CASES_ARRAY(cash_er_tf_const, er_case) = FALSE
' 			Else
' 				ER_CASES_ARRAY(cash_er_tf_const, er_case) = TRUE
' 			End If
' 			If snap_er_code = "-" OR snap_er_code = "_" OR snap_er_code = " " Then
' 				ER_CASES_ARRAY(snap_er_tf_const, er_case) = FALSE
' 				ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
' 			Else
' 				ER_CASES_ARRAY(snap_er_tf_const, er_case) = TRUE
' 				ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE
' 				' MsgBox "Case Number - " & ER_CASES_ARRAY(case_nbr_const, er_case) & vbCr & "Cash REVW - " & cash_grh_er_code & vbCr & "SNAP REVW - " & snap_er_code & vbCr & "CAF DATE - " & revs_caf_date & vbCr & "INTERVIEW DATE - " & revs_intv_date & vbCr & ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case)
' 			End If
' 			If revs_caf_date <> "__ __ __" Then ER_CASES_ARRAY(caf_date_const, er_case) = replace(revs_caf_date, " ", "/")
' 			If revs_intv_date <> "__ __ __" Then
' 				ER_CASES_ARRAY(intv_date_const, er_case) = replace(revs_intv_date, " ", "/")
' 				ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
' 			End If
' 			' MsgBox "Case Number - " & ER_CASES_ARRAY(case_nbr_const, er_case) & vbCr & "Cash REVW - " & cash_grh_er_code & vbCr & "SNAP REVW - " & snap_er_code & vbCr & "CAF DATE - " & revs_caf_date & vbCr & "INTERVIEW DATE - " & revs_intv_date & vbCr & ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case)
' 		ELSE
' 			ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = "FAILED"
' 		End If
' 		Call clear_line_of_text(20, 33)
' 	End If
' Next
'reat rept revw for caf date and interview date
'read rept revs for cash er t/f
'read rept revs for snap er t/f

Call Back_to_self
'IF cash er t then read cash program active
For er_case = 0 To UBound(ER_CASES_ARRAY, 2)
	' If ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE Then
	' 	If ER_CASES_ARRAY(cash_er_tf_const, er_case) = TRUE Then
	' 		MAXIS_case_number = ER_CASES_ARRAY(case_nbr_const, er_case)
	'
	' 		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)
	' 		If mfip_case = TRUE Then ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE
	'
	' 		If mfip_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", MFIP"
	' 		If ga_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", GA"
	' 		If msa_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", MSA"
	' 		If grh_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", GRH"
	' 		If left(ER_CASES_ARRAY(cash_prog_const, er_case), 2) = ", " Then ER_CASES_ARRAY(cash_prog_const, er_case) = right(ER_CASES_ARRAY(cash_prog_const, er_case), len(ER_CASES_ARRAY(cash_prog_const, er_case)) - 2)
	'
	' 		Call Back_to_self
	' 	End If
	' End If
	' If ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = TRUE Then
	' 	If ER_CASES_ARRAY(cash_er_tf_const, er_case) = TRUE Then
			MAXIS_case_number = ER_CASES_ARRAY(case_nbr_const, er_case)

			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
			' If mfip_case = TRUE Then ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) = FALSE

			If mfip_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", MFIP"
			If ga_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", GA"
			If msa_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", MSA"
			If grh_case = TRUE THen ER_CASES_ARRAY(cash_prog_const, er_case) = ER_CASES_ARRAY(cash_prog_const, er_case) & ", GRH"
			If left(ER_CASES_ARRAY(cash_prog_const, er_case), 2) = ", " Then ER_CASES_ARRAY(cash_prog_const, er_case) = right(ER_CASES_ARRAY(cash_prog_const, er_case), len(ER_CASES_ARRAY(cash_prog_const, er_case)) - 2)

			Call Back_to_self
	' 	End If
	' End If
	row = ER_CASES_ARRAY(excel_row_const, er_case)
	ObjExcel.Cells(row, 9).Value = ER_CASES_ARRAY(poss_waived_intv_tf_const, er_case) & ""
	ObjExcel.Cells(row, 10).Value = ER_CASES_ARRAY(cash_prog_const, er_case)
Next
'If cash er is false then possible waiver is true
'if cash prog is not MFIP then possible waiver is TRUE
'output all the case with possible waiver into the RENEWAL file - true and include snap er t/f, cash er t/f, cash prog
script_end_procedure("done now")
