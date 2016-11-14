msgbox "If you are seeing this message please notify Charles Potter at Charles.D.Potter@state.mn.us"

'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "BULK - REPT-ARST LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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

'DIALOGS-------------------------------------------------------
BeginDialog REPT_ARST_dialog, 0, 0, 276, 135, "REPT ARST Dialog"
  CheckBox 10, 15, 60, 10, "Cash", cash_check
  CheckBox 10, 30, 60, 10, "EMER", EMER_check
  CheckBox 10, 45, 60, 10, "HC", HC_check
  CheckBox 10, 60, 60, 10, "SNAP", SNAP_check
  CheckBox 10, 75, 60, 10, "GRH", GRH_check
  CheckBox 100, 15, 70, 10, "Address Changes", address_changes_check
  CheckBox 100, 30, 70, 10, "Active", active_check
  CheckBox 100, 45, 70, 10, "Pending total", pending_check
  CheckBox 100, 60, 70, 10, "REIN", REIN_check
  CheckBox 180, 15, 85, 10, "Pending < 31 days", pending_under_31_check
  CheckBox 180, 30, 85, 10, "Pending 31-45 days", pending_31_to_45_check
  CheckBox 180, 45, 85, 10, "Pending 46-60 days", pending_46_to_60_check
  CheckBox 180, 60, 85, 10, "Pending > 60 days", pending_over_60_check
  EditBox 55, 95, 25, 15, MAXIS_footer_month
  EditBox 55, 115, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 215, 95, 50, 15
    CancelButton 215, 115, 50, 15
    PushButton 105, 100, 95, 10, "All Possible Data", all_possible_data_button
    PushButton 105, 110, 95, 10, "SNAP Pending Info", SNAP_pending_info_button
    PushButton 105, 120, 95, 10, "Active, Pending, and REINs", ACTV_pending_REIN_button
  GroupBox 5, 5, 80, 85, "Program"
  GroupBox 95, 5, 175, 70, "Number to pull"
  GroupBox 95, 85, 110, 50, "Common options"
  Text 5, 100, 50, 10, "Footer month:"
  Text 5, 120, 40, 10, "Footer year:"
EndDialog

'DEFINING VARIABLES----------------------------------------------------------------------------------------------------
excel_row = 3 'this is the row the workers will start on the spreadsheet
MAXIS_footer_month = datepart("m", date) & ""		'Footer month defaults to this month
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month	'In case this month is a single digit month
MAXIS_footer_year = right(datepart("yyyy", date), 2)	'Footer year is the right two digits of the current year

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Showing dialog
Do
	Dialog REPT_ARST_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = all_possible_data_button then
		cash_check = checked
		EMER_check = checked
		HC_check = checked
		SNAP_check = checked
		GRH_check = checked
		address_changes_check = checked
		active_check = checked
		pending_check = checked
		REIN_check = checked
		pending_under_31_check = checked
		pending_31_to_45_check = checked
		pending_46_to_60_check = checked
		pending_over_60_check = checked
	End if
	If buttonpressed = SNAP_pending_info_button then
		cash_check = unchecked
		EMER_check = unchecked
		HC_check = unchecked
		SNAP_check = checked
		GRH_check = unchecked
		address_changes_check = unchecked
		active_check = unchecked
		pending_check = checked
		REIN_check = unchecked
		pending_under_31_check = checked
		pending_31_to_45_check = checked
		pending_46_to_60_check = checked
		pending_over_60_check = checked
	End if
	If buttonpressed = ACTV_pending_REIN_button then
		cash_check = checked
		EMER_check = checked
		HC_check = checked
		SNAP_check = checked
		GRH_check = checked
		address_changes_check = unchecked
		active_check = checked
		pending_check = checked
		REIN_check = checked
		pending_under_31_check = unchecked
		pending_31_to_45_check = unchecked
		pending_46_to_60_check = unchecked
		pending_over_60_check = unchecked
	End if
Loop until buttonpressed = OK

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking MAXIS
Call check_for_MAXIS(True)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker and name, using row 2 because row 1 will contain headers for programs
ObjExcel.Cells(2, 1).Value = "WORKER NUMBER"
objExcel.Cells(2, 1).Font.Bold = TRUE
ObjExcel.Cells(2, 2).Value = "WORKER NAME"
objExcel.Cells(2, 2).Font.Bold = TRUE

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 3 'Starting with 3 because cols 1-2 are already used

'creates a gather_cases_pending veriable as TRUE if any of the pending options is checked to get case data correctly 
If pending_check = checked OR pending_under_31_check = checked OR pending_31_to_45_check = checked OR pending_46_to_60_check = checked OR pending_over_60_check = checked THEN gather_cases_pending = TRUE 

case_header_col = col_to_use		'Sets the header to be used later for merging Cells

'Sets the Excel Sheet up to document CASE information - headers and variable information
ObjExcel.Cells(2, col_to_use).Value = "Total"
objExcel.Cells(2, col_to_use).Font.Bold = TRUE
case_total_col = col_to_use
objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
col_to_use = col_to_use + 1
case_total_letter_col = convert_digit_to_excel_column(case_total_col)

If active_check = checked then
	ObjExcel.Cells(2, col_to_use).Value = "ACTV"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	case_actv_col = col_to_use
	objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
	col_to_use = col_to_use + 1
	case_actv_letter_col = convert_digit_to_excel_column(case_actv_col)
End if
If REIN_check = checked then
	ObjExcel.Cells(2, col_to_use).Value = "REIN"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	case_REIN_col = col_to_use
	objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
	col_to_use = col_to_use + 1
	case_REIN_letter_col = convert_digit_to_excel_column(case_REIN_col)
End if
If gather_cases_pending = TRUE then
	ObjExcel.Cells(2, col_to_use).Value = "PND2"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	case_pnd2_col = col_to_use
	objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
	col_to_use = col_to_use + 1
	case_pnd2_letter_col = convert_digit_to_excel_column(case_pnd2_col)

	ObjExcel.Cells(2, col_to_use).Value = "PND1"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	case_pnd1_col = col_to_use
	objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
	col_to_use = col_to_use + 1
	case_pnd1_letter_col = convert_digit_to_excel_column(case_pnd1_col)
End if

'Header cell
ObjExcel.Cells(1, case_header_col).Value = "CASES"
objExcel.Cells(1, case_header_col).Font.Bold = TRUE

'Merging header cell. Uses col_to_use - 1 because we already moved on to the next column.
ObjExcel.Range(ObjExcel.Cells(1, case_header_col), ObjExcel.Cells(1, col_to_use - 1)).Merge

'Centering the cell
objExcel.Cells(1, case_header_col).HorizontalAlignment = -4108

'Headers and variable declaration for cash
If cash_check = checked then
	cash_header_col = col_to_use 'will use this later to merge cells

	If active_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "ACTV"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_actv_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_actv_letter_col = convert_digit_to_excel_column(cash_actv_col)
	End if
	If REIN_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "REIN"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_REIN_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_REIN_letter_col = convert_digit_to_excel_column(cash_REIN_col)
	End if
	If pending_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "total"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_pending_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_pending_letter_col = convert_digit_to_excel_column(cash_pending_col)
	End if
	If pending_under_31_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "<31"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_pending_under_31_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_pending_under_31_letter_col = convert_digit_to_excel_column(cash_pending_under_31_col)
	End if
	If pending_31_to_45_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "31-45"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_pending_31_to_45_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_pending_31_to_45_letter_col = convert_digit_to_excel_column(cash_pending_31_to_45_col)
	End if
	If pending_46_to_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "46-60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_pending_46_to_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_pending_46_to_60_letter_col = convert_digit_to_excel_column(cash_pending_46_to_60_col)
	End if
	If pending_over_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & ">60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		cash_pending_over_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		cash_pending_over_60_letter_col = convert_digit_to_excel_column(cash_pending_over_60_col)
	End if

	'Header cell
	ObjExcel.Cells(1, cash_header_col).Value = "CASH"
	objExcel.Cells(1, cash_header_col).Font.Bold = TRUE

	'Merging header cell. Uses col_to_use - 1 because we already moved on to the next column.
	ObjExcel.Range(ObjExcel.Cells(1, cash_header_col), ObjExcel.Cells(1, col_to_use - 1)).Merge

	'Centering the cell
	objExcel.Cells(1, cash_header_col).HorizontalAlignment = -4108
End if

'Headers and variable declaration for SNAP
If SNAP_check = checked then
	SNAP_header_col = col_to_use 'will use this later to merge cells

	If active_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "ACTV"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_actv_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_actv_letter_col = convert_digit_to_excel_column(SNAP_actv_col)
	End if
	If REIN_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "REIN"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_REIN_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_REIN_letter_col = convert_digit_to_excel_column(SNAP_REIN_col)
	End if
	If pending_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "total"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_pending_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_pending_letter_col = convert_digit_to_excel_column(SNAP_pending_col)
	End if
	If pending_under_31_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "<31"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_pending_under_31_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_pending_under_31_letter_col = convert_digit_to_excel_column(SNAP_pending_under_31_col)
	End if
	If pending_31_to_45_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "31-45"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_pending_31_to_45_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_pending_31_to_45_letter_col = convert_digit_to_excel_column(SNAP_pending_31_to_45_col)
	End if
	If pending_46_to_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "46-60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_pending_46_to_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_pending_46_to_60_letter_col = convert_digit_to_excel_column(SNAP_pending_46_to_60_col)
	End if
	If pending_over_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & ">60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		SNAP_pending_over_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		SNAP_pending_over_60_letter_col = convert_digit_to_excel_column(SNAP_pending_over_60_col)
	End if

	'Header cell
	ObjExcel.Cells(1, SNAP_header_col).Value = "SNAP"
	objExcel.Cells(1, SNAP_header_col).Font.Bold = TRUE

	'Merging header cell. Uses col_to_use - 1 because we already moved on to the next column.
	ObjExcel.Range(ObjExcel.Cells(1, SNAP_header_col), ObjExcel.Cells(1, col_to_use - 1)).Merge

	'Centering the cell
	objExcel.Cells(1, SNAP_header_col).HorizontalAlignment = -4108
End if

'Headers and variable declaration for HC
If HC_check = checked then
	HC_header_col = col_to_use 'will use this later to merge cells

	If active_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "ACTV"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_actv_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_actv_letter_col = convert_digit_to_excel_column(HC_actv_col)
	End if
	If REIN_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "REIN"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_REIN_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_REIN_letter_col = convert_digit_to_excel_column(HC_REIN_col)
	End if
	If pending_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "total"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_pending_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_pending_letter_col = convert_digit_to_excel_column(HC_pending_col)
	End if
	If pending_under_31_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "<31"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_pending_under_31_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_pending_under_31_letter_col = convert_digit_to_excel_column(HC_pending_under_31_col)
	End if
	If pending_31_to_45_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "31-45"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_pending_31_to_45_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_pending_31_to_45_letter_col = convert_digit_to_excel_column(HC_pending_31_to_45_col)
	End if
	If pending_46_to_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "46-60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_pending_46_to_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_pending_46_to_60_letter_col = convert_digit_to_excel_column(HC_pending_46_to_60_col)
	End if
	If pending_over_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & ">60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		HC_pending_over_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		HC_pending_over_60_letter_col = convert_digit_to_excel_column(HC_pending_over_60_col)
	End if

	'Header cell
	ObjExcel.Cells(1, HC_header_col).Value = "HC"
	objExcel.Cells(1, HC_header_col).Font.Bold = TRUE

	'Merging header cell. Uses col_to_use - 1 because we already moved on to the next column.
	ObjExcel.Range(ObjExcel.Cells(1, HC_header_col), ObjExcel.Cells(1, col_to_use - 1)).Merge

	'Centering the cell
	objExcel.Cells(1, HC_header_col).HorizontalAlignment = -4108
End if

'Headers and variable declaration for GRH
If GRH_check = checked then
	GRH_header_col = col_to_use 'will use this later to merge cells

	If active_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "ACTV"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_actv_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_actv_letter_col = convert_digit_to_excel_column(GRH_actv_col)
	End if
	If REIN_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "REIN"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_REIN_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_REIN_letter_col = convert_digit_to_excel_column(GRH_REIN_col)
	End if
	If pending_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "total"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_pending_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_pending_letter_col = convert_digit_to_excel_column(GRH_pending_col)
	End if
	If pending_under_31_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "<31"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_pending_under_31_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_pending_under_31_letter_col = convert_digit_to_excel_column(GRH_pending_under_31_col)
	End if
	If pending_31_to_45_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "31-45"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_pending_31_to_45_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_pending_31_to_45_letter_col = convert_digit_to_excel_column(GRH_pending_31_to_45_col)
	End if
	If pending_46_to_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "46-60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_pending_46_to_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_pending_46_to_60_letter_col = convert_digit_to_excel_column(GRH_pending_46_to_60_col)
	End if
	If pending_over_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & ">60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		GRH_pending_over_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		GRH_pending_over_60_letter_col = convert_digit_to_excel_column(GRH_pending_over_60_col)
	End if

	'Header cell
	ObjExcel.Cells(1, GRH_header_col).Value = "GRH"
	objExcel.Cells(1, GRH_header_col).Font.Bold = TRUE

	'Merging header cell. Uses col_to_use - 1 because we already moved on to the next column.
	ObjExcel.Range(ObjExcel.Cells(1, GRH_header_col), ObjExcel.Cells(1, col_to_use - 1)).Merge

	'Centering the cell
	objExcel.Cells(1, GRH_header_col).HorizontalAlignment = -4108
End if

'Headers and variable declaration for EMER
If EMER_check = checked then
	EMER_header_col = col_to_use 'will use this later to merge cells

	If active_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "ACTV"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_actv_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_actv_letter_col = convert_digit_to_excel_column(EMER_actv_col)
	End if
	If REIN_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "REIN"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_REIN_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_REIN_letter_col = convert_digit_to_excel_column(EMER_REIN_col)
	End if
	If pending_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "total"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_pending_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_pending_letter_col = convert_digit_to_excel_column(EMER_pending_col)
	End if
	If pending_under_31_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbnewline & "<31"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_pending_under_31_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_pending_under_31_letter_col = convert_digit_to_excel_column(EMER_pending_under_31_col)
	End if
	If pending_31_to_45_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "31-45"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_pending_31_to_45_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_pending_31_to_45_letter_col = convert_digit_to_excel_column(EMER_pending_31_to_45_col)
	End if
	If pending_46_to_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & "46-60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_pending_46_to_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_pending_46_to_60_letter_col = convert_digit_to_excel_column(EMER_pending_46_to_60_col)
	End if
	If pending_over_60_check = checked then
		ObjExcel.Cells(2, col_to_use).Value = "pend" & vbNewLine & ">60"
		objExcel.Cells(2, col_to_use).Font.Bold = TRUE
		EMER_pending_over_60_col = col_to_use
		objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
		col_to_use = col_to_use + 1
		EMER_pending_over_60_letter_col = convert_digit_to_excel_column(EMER_pending_over_60_col)
	End if

	'Header cell
	ObjExcel.Cells(1, EMER_header_col).Value = "EMER"
	objExcel.Cells(1, EMER_header_col).Font.Bold = TRUE

	'Merging header cell. Uses col_to_use - 1 because we already moved on to the next column.
	ObjExcel.Range(ObjExcel.Cells(1, EMER_header_col), ObjExcel.Cells(1, col_to_use - 1)).Merge

	'Centering the cell
	objExcel.Cells(1, EMER_header_col).HorizontalAlignment = -4108
End if

'Address changes info
If address_changes_check = checked then
	ObjExcel.Cells(2, col_to_use).Value = "address" & vbNewLine & "changes"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	address_changes_col = col_to_use
	objExcel.Cells(2, col_to_use).HorizontalAlignment = -4108
	col_to_use = col_to_use + 1
	address_changes_letter_col = convert_digit_to_excel_column(address_changes_col)
End if

'Grabbing all workers from county
call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

'Getting to SELF (need to reset the footer month)
back_to_self

'Resetting the footer month
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
transmit

'Getting to REPT/ARST
call navigate_to_MAXIS_screen("rept", "arst")

'Reading the accumulations timestamp to be used later when we provide stats to the user
EMReadScreen accumulations_timestamp, 30, 19, 40
accumulations_timestamp = trim(accumulations_timestamp)

'Actually collecting the info
For each worker_number in worker_array
	EMWriteScreen worker_number, 3, 27	'Putting worker number onto ARST
	transmit

	'The following will determine if the worker is active or inactive. Inactive workers will show a 0 for all numbers, and will not be entered into the spreadsheet
	EMReadScreen total_cases, 			 9, 5, 47
	EMReadScreen cases_actv, 			 9, 6, 47 
	EMReadScreen cases_rein, 			 9, 7, 47
	EMReadScreen cases_pnd2, 			 9, 8, 47
	EMReadScreen cases_pnd1, 			 9, 9, 47
	EMReadScreen CAF_I_APPL_taken, 	 	 9, 12, 47
	EMReadScreen CAF_II_APPL_taken, 	 9, 13, 47
	EMReadScreen cases_auto_denied, 	 9, 14, 47
	EMReadScreen address_changes_count,  9, 15, 47	'Reading address changes, which will be used later
	EMReadScreen ASET_versions_approved, 9, 16, 47

	'Deciding if the case is active based on the above responses.
	If trim(total_cases) <> "0" or _
	trim(CAF_I_APPL_taken) <> "0" or _
	trim(CAF_II_APPL_taken) <> "0" or _
	trim(cases_auto_denied) <> "0" or _
	trim(address_changes_count) <> "0" or _
	trim(ASET_versions_approved) <> "0" then
		worker_is_active = True
	Else
		worker_is_active = false
	End if

	If worker_is_active = True then

		PF8							'Navigating to next screen

		'Reading cash info
		EMReadScreen cash_active_count, 7, 9, 21
		EMReadScreen cash_REIN_count, 7, 9, 29
		EMReadScreen cash_pending_count, 7, 9, 37
		EMReadScreen cash_pending_under_31_count, 7, 9, 45
		EMReadScreen cash_pending_31_to_45_count, 7, 9, 53
		EMReadScreen cash_pending_46_to_60_count, 7, 9, 61
		EMReadScreen cash_pending_over_60_count, 7, 9, 69

		'Navigating to next screen
		PF8

		'Reading GRH info
		EMReadScreen GRH_active_count, 7, 8, 21
		EMReadScreen GRH_REIN_count, 7, 8, 29
		EMReadScreen GRH_pending_count, 7, 8, 37
		EMReadScreen GRH_pending_under_31_count, 7, 8, 45
		EMReadScreen GRH_pending_31_to_45_count, 7, 8, 53
		EMReadScreen GRH_pending_46_to_60_count, 7, 8, 61
		EMReadScreen GRH_pending_over_60_count, 7, 8, 69

		'Reading EMER info
		EMReadScreen EMER_active_count, 7, 9, 21
		EMReadScreen EMER_REIN_count, 7, 9, 29
		EMReadScreen EMER_pending_count, 7, 9, 37
		EMReadScreen EMER_pending_under_31_count, 7, 9, 45
		EMReadScreen EMER_pending_31_to_45_count, 7, 9, 53
		EMReadScreen EMER_pending_46_to_60_count, 7, 9, 61
		EMReadScreen EMER_pending_over_60_count, 7, 9, 69

		'Reading SNAP info
		EMReadScreen SNAP_active_count, 7, 14, 21
		EMReadScreen SNAP_REIN_count, 7, 14, 29
		EMReadScreen SNAP_pending_count, 7, 14, 37
		EMReadScreen SNAP_pending_under_31_count, 7, 14, 45
		EMReadScreen SNAP_pending_31_to_45_count, 7, 14, 53
		EMReadScreen SNAP_pending_46_to_60_count, 7, 14, 61
		EMReadScreen SNAP_pending_over_60_count, 7, 14, 69

		'Navigatingto next screen
		PF8

		'Reading HC info
		EMReadScreen HC_active_count, 7, 8, 21
		EMReadScreen HC_REIN_count, 7, 8, 29
		EMReadScreen HC_pending_count, 7, 8, 37
		EMReadScreen HC_pending_under_31_count, 7, 8, 45
		EMReadScreen HC_pending_31_to_45_count, 7, 8, 53
		EMReadScreen HC_pending_46_to_60_count, 7, 8, 61
		EMReadScreen HC_pending_over_60_count, 7, 8, 69

		'Getting back to first screen
		PF7
		PF7
		PF7

		'Writing to Excel
		ObjExcel.Cells(excel_row, 1).Value = worker_number
		ObjExcel.Cells(excel_row, 2).Value = worker_name

		ObjExcel.Cells(excel_row, case_total_col).Value = total_cases
		If active_check = checked then ObjExcel.Cells(excel_row, case_actv_col).Value = cases_actv
		If REIN_check = checked then ObjExcel.Cells(excel_row, case_REIN_col).Value = cases_rein
		If gather_cases_pending = TRUE then 
			ObjExcel.Cells(excel_row, case_pnd2_col).Value = cases_pnd2
			ObjExcel.Cells(excel_row, case_pnd1_col).Value = cases_pnd1
		End If 
		
		'Cash info to Excel
		If cash_check = checked then
			If active_check = checked then ObjExcel.Cells(excel_row, cash_actv_col).Value = cash_active_count
			If REIN_check = checked then ObjExcel.Cells(excel_row, cash_REIN_col).Value = cash_REIN_count
			If pending_check = checked then ObjExcel.Cells(excel_row, cash_pending_col).Value = cash_pending_count
			If pending_under_31_check = checked then ObjExcel.Cells(excel_row, cash_pending_under_31_col).Value = cash_pending_under_31_count
			If pending_31_to_45_check = checked then ObjExcel.Cells(excel_row, cash_pending_31_to_45_col).Value = cash_pending_31_to_45_count
			If pending_46_to_60_check = checked then ObjExcel.Cells(excel_row, cash_pending_46_to_60_col).Value = cash_pending_46_to_60_count
			If pending_over_60_check = checked then ObjExcel.Cells(excel_row, cash_pending_over_60_col).Value = cash_pending_over_60_count
		End if

		'SNAP info to Excel
		If SNAP_check = checked then
			If active_check = checked then ObjExcel.Cells(excel_row, SNAP_actv_col).Value = SNAP_active_count
			If REIN_check = checked then ObjExcel.Cells(excel_row, SNAP_REIN_col).Value = SNAP_REIN_count
			If pending_check = checked then ObjExcel.Cells(excel_row, SNAP_pending_col).Value = SNAP_pending_count
			If pending_under_31_check = checked then ObjExcel.Cells(excel_row, SNAP_pending_under_31_col).Value = SNAP_pending_under_31_count
			If pending_31_to_45_check = checked then ObjExcel.Cells(excel_row, SNAP_pending_31_to_45_col).Value = SNAP_pending_31_to_45_count
			If pending_46_to_60_check = checked then ObjExcel.Cells(excel_row, SNAP_pending_46_to_60_col).Value = SNAP_pending_46_to_60_count
			If pending_over_60_check = checked then ObjExcel.Cells(excel_row, SNAP_pending_over_60_col).Value = SNAP_pending_over_60_count
		End if

		'HC info to Excel
		If HC_check = checked then
			If active_check = checked then ObjExcel.Cells(excel_row, HC_actv_col).Value = HC_active_count
			If REIN_check = checked then ObjExcel.Cells(excel_row, HC_REIN_col).Value = HC_REIN_count
			If pending_check = checked then ObjExcel.Cells(excel_row, HC_pending_col).Value = HC_pending_count
			If pending_under_31_check = checked then ObjExcel.Cells(excel_row, HC_pending_under_31_col).Value = HC_pending_under_31_count
			If pending_31_to_45_check = checked then ObjExcel.Cells(excel_row, HC_pending_31_to_45_col).Value = HC_pending_31_to_45_count
			If pending_46_to_60_check = checked then ObjExcel.Cells(excel_row, HC_pending_46_to_60_col).Value = HC_pending_46_to_60_count
			If pending_over_60_check = checked then ObjExcel.Cells(excel_row, HC_pending_over_60_col).Value = HC_pending_over_60_count
		End if

		'GRH info to Excel
		If GRH_check = checked then
			If active_check = checked then ObjExcel.Cells(excel_row, GRH_actv_col).Value = GRH_active_count
			If REIN_check = checked then ObjExcel.Cells(excel_row, GRH_REIN_col).Value = GRH_REIN_count
			If pending_check = checked then ObjExcel.Cells(excel_row, GRH_pending_col).Value = GRH_pending_count
			If pending_under_31_check = checked then ObjExcel.Cells(excel_row, GRH_pending_under_31_col).Value = GRH_pending_under_31_count
			If pending_31_to_45_check = checked then ObjExcel.Cells(excel_row, GRH_pending_31_to_45_col).Value = GRH_pending_31_to_45_count
			If pending_46_to_60_check = checked then ObjExcel.Cells(excel_row, GRH_pending_46_to_60_col).Value = GRH_pending_46_to_60_count
			If pending_over_60_check = checked then ObjExcel.Cells(excel_row, GRH_pending_over_60_col).Value = GRH_pending_over_60_count
		End if

		'EMER info to Excel
		If EMER_check = checked then
			If active_check = checked then ObjExcel.Cells(excel_row, EMER_actv_col).Value = EMER_active_count
			If REIN_check = checked then ObjExcel.Cells(excel_row, EMER_REIN_col).Value = EMER_REIN_count
			If pending_check = checked then ObjExcel.Cells(excel_row, EMER_pending_col).Value = EMER_pending_count
			If pending_under_31_check = checked then ObjExcel.Cells(excel_row, EMER_pending_under_31_col).Value = EMER_pending_under_31_count
			If pending_31_to_45_check = checked then ObjExcel.Cells(excel_row, EMER_pending_31_to_45_col).Value = EMER_pending_31_to_45_count
			If pending_46_to_60_check = checked then ObjExcel.Cells(excel_row, EMER_pending_46_to_60_col).Value = EMER_pending_46_to_60_count
			If pending_over_60_check = checked then ObjExcel.Cells(excel_row, EMER_pending_over_60_col).Value = EMER_pending_over_60_count
		End if

		'Address changes info to Excel
		If address_changes_check = checked then ObjExcel.Cells(excel_row, address_changes_col).Value = address_changes_count

		excel_row = excel_row + 1

	End if
Next

'Now it gets the worker names for each worker on the spreadsheet

'Resets excel_row, so as to read back the names
excel_row = 3

'Navigates to REPT/USER
call navigate_to_MAXIS_screen("REPT", "USER")

'This do...loop will read the worker x1, navigate to REPT/USER, and get the worker's name
Do
	x1_number_to_read = ObjExcel.Cells(excel_row, 1).Value		'Assigns the x1 in the spreadsheet to a variable
	EMWriteScreen x1_number_to_read, 21, 12					'Writes that variable into the PW number field in MAXIS
	transmit
	EMWriteScreen "x", 7, 3								'Selects the "view more" option
	transmit
	EMReadScreen worker_name_from_USER, 40, 7, 27				'Reads the entire worker name
	ObjExcel.Cells(excel_row, 2).Value = trim(worker_name_from_USER)	'Puts the worker_name in the spreadsheet next to the x1 number
	excel_row = excel_row + 1							'We'll need to look at the next row
	PF3											'Gets out of this screen
Loop until ObjExcel.Cells(excel_row, 1).Value = ""	'Exit loop when it's blank


col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'Declaring here before the following if...then statements

'Query date/time/runtime info
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"			'Goes back one, as this is on the next row
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"		'Goes back one, as this is on the next row
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time
ObjExcel.Cells(3, col_to_use - 1).Value = "MAXIS accumulation timestamp:"	'Goes back one, as this is on the next row
objExcel.Cells(3, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(3, col_to_use).Value = accumulations_timestamp

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Next

'This bit freezes the top 2 rows for scrolling ease of use
ObjExcel.ActiveSheet.Range("A3").Select
objExcel.ActiveWindow.FreezePanes = True

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! The statistics have loaded.")
