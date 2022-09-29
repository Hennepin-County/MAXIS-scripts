'Required for statistical purposes==========================================================================================
name_of_script = "MISC - CHECK UNASSIGNED EXPEDITED CASES.vbs"
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

function find_q_flow_population(basket_number, suggested_population)
	If basket_number = "X127EF8" then suggested_population = "1800"
	If basket_number = "X127EF9" then suggested_population = "1800"
	If basket_number = "X127EG9" then suggested_population = "1800"
	If basket_number = "X127EG0" then suggested_population = "1800"

	If basket_number = "X127ED8" then suggested_population = "Adults"
	If basket_number = "X127EE1" then suggested_population = "Adults"
	If basket_number = "X127EE2" then suggested_population = "Adults"
	If basket_number = "X127EE3" then suggested_population = "Adults"
	If basket_number = "X127EE4" then suggested_population = "Adults"
	If basket_number = "X127EE5" then suggested_population = "Adults"
	If basket_number = "X127EE6" then suggested_population = "Adults"
	If basket_number = "X127EE7" then suggested_population = "Adults"
	If basket_number = "X127EG4" then suggested_population = "Adults"
	If basket_number = "X127EH8" then suggested_population = "Adults"
	If basket_number = "X127EJ1" then suggested_population = "Adults"
	If basket_number = "X127EL1" then suggested_population = "Adults"
	If basket_number = "X127EL2" then suggested_population = "Adults"
	If basket_number = "X127EL3" then suggested_population = "Adults"
	If basket_number = "X127EL4" then suggested_population = "Adults"
	If basket_number = "X127EL5" then suggested_population = "Adults"
	If basket_number = "X127EL6" then suggested_population = "Adults"
	If basket_number = "X127EL7" then suggested_population = "Adults"
	If basket_number = "X127EL8" then suggested_population = "Adults"
	If basket_number = "X127EL9" then suggested_population = "Adults"
	If basket_number = "X127EN1" then suggested_population = "Adults"
	If basket_number = "X127EN2" then suggested_population = "Adults"
	If basket_number = "X127EN3" then suggested_population = "Adults"
	If basket_number = "X127EN4" then suggested_population = "Adults"
	If basket_number = "X127EN5" then suggested_population = "Adults"
	If basket_number = "X127EN7" then suggested_population = "Adults"
	If basket_number = "X127EP6" then suggested_population = "Adults"
	If basket_number = "X127EP7" then suggested_population = "Adults"
	If basket_number = "X127EP8" then suggested_population = "Adults"
	If basket_number = "X127EQ1" then suggested_population = "Adults"
	If basket_number = "X127EQ3" then suggested_population = "Adults"
	If basket_number = "X127EQ4" then suggested_population = "Adults"
	If basket_number = "X127EQ5" then suggested_population = "Adults"
	If basket_number = "X127EQ8" then suggested_population = "Adults"
	If basket_number = "X127EQ9" then suggested_population = "Adults"
	If basket_number = "X127EX1" then suggested_population = "Adults"
	If basket_number = "X127EX2" then suggested_population = "Adults"
	If basket_number = "X127EX3" then suggested_population = "Adults"
	If basket_number = "X127EX7" then suggested_population = "Adults"
	If basket_number = "X127EX8" then suggested_population = "Adults"
	If basket_number = "X127EX9" then suggested_population = "Adults"
	If basket_number = "X127F3D" then suggested_population = "Adults"
	If basket_number = "X127F3P" then suggested_population = "Adults"   'MA-EPD Adults Basket

	If basket_number = "X127FE7" then suggested_population = "DWP"
	If basket_number = "X127FE8" then suggested_population = "DWP"
	If basket_number = "X127FE9" then suggested_population = "DWP"

	If basket_number = "X127EP8" then suggested_population = "EGA"
	If basket_number = "X127EQ2" then suggested_population = "EGA"

	If basket_number = "X127ES1" then suggested_population = "Families"
	If basket_number = "X127ES2" then suggested_population = "Families"
	If basket_number = "X127ES3" then suggested_population = "Families"
	If basket_number = "X127ES4" then suggested_population = "Families"
	If basket_number = "X127ES5" then suggested_population = "Families"
	If basket_number = "X127ES6" then suggested_population = "Families"
	If basket_number = "X127ES7" then suggested_population = "Families"
	If basket_number = "X127ES8" then suggested_population = "Families"
	If basket_number = "X127ES9" then suggested_population = "Families"
	If basket_number = "X127ET1" then suggested_population = "Families"
	If basket_number = "X127ET2" then suggested_population = "Families"
	If basket_number = "X127ET3" then suggested_population = "Families"
	If basket_number = "X127ET4" then suggested_population = "Families"
	If basket_number = "X127ET5" then suggested_population = "Families"
	If basket_number = "X127ET6" then suggested_population = "Families"
	If basket_number = "X127ET7" then suggested_population = "Families"
	If basket_number = "X127ET8" then suggested_population = "Families"
	If basket_number = "X127ET9" then suggested_population = "Families"
	If basket_number = "X127F4E" then suggested_population = "Families"
	If basket_number = "X127F3H" then suggested_population = "Families"
	If basket_number = "X127FB7" then suggested_population = "Families"
	If basket_number = "X127EZ1" then suggested_population = "Families"
	If basket_number = "X127EZ3" then suggested_population = "Families"
	If basket_number = "X127EZ4" then suggested_population = "Families"
	If basket_number = "X127EZ6" then suggested_population = "Families"
	If basket_number = "X127EZ7" then suggested_population = "Families"
	If basket_number = "X127EZ8" then suggested_population = "Families"
	If basket_number = "X127F3K" then suggested_population = "Families"  'MA-EPD FAD Basket

	If basket_number = "X127EZ2" then suggested_population = "FAD GRH"

	If basket_number = "X127EG5" then suggested_population = "Housing Supports"
	If basket_number = "X127FG3" then suggested_population = "Housing Supports"
	If basket_number = "X127EH2" then suggested_population = "Housing Supports"
	If basket_number = "X127EJ4" then suggested_population = "Housing Supports"
	If basket_number = "X127EJ7" then suggested_population = "Housing Supports"
	If basket_number = "X127EK5" then suggested_population = "Housing Supports"
	If basket_number = "X127EM1" then suggested_population = "Housing Supports"
	If basket_number = "X127EM8" then suggested_population = "Housing Supports"
	If basket_number = "X127EP4" then suggested_population = "Housing Supports"

	If basket_number = "X127EH1" then suggested_population = "LTC+"
	If basket_number = "X127EH3" then suggested_population = "LTC+"
	If basket_number = "X127EH4" then suggested_population = "LTC+"
	If basket_number = "X127EH5" then suggested_population = "LTC+"
	If basket_number = "X127EH6" then suggested_population = "LTC+"
	If basket_number = "X127EH7" then suggested_population = "LTC+"
	If basket_number = "X127EJ8" then suggested_population = "LTC+"
	If basket_number = "X127EK1" then suggested_population = "LTC+"
	If basket_number = "X127EK2" then suggested_population = "LTC+"
	If basket_number = "X127EK3" then suggested_population = "LTC+"
	If basket_number = "X127EK4" then suggested_population = "LTC+"
	If basket_number = "X127EK6" then suggested_population = "LTC+"
	If basket_number = "X127EK7" then suggested_population = "LTC+"
	If basket_number = "X127EK8" then suggested_population = "LTC+"
	If basket_number = "X127EK9" then suggested_population = "LTC+"
	If basket_number = "X127EM9" then suggested_population = "LTC+"
	If basket_number = "X127EN6" then suggested_population = "LTC+"
	If basket_number = "X127EP5" then suggested_population = "LTC+"
	If basket_number = "X127EP9" then suggested_population = "LTC+"
	If basket_number = "X127EZ5" then suggested_population = "LTC+"
	If basket_number = "X127F3F" then suggested_population = "LTC+"
	If basket_number = "X127FE5" then suggested_population = "LTC+"
	If basket_number = "X127FH4" then suggested_population = "LTC+"
	If basket_number = "X127FH5" then suggested_population = "LTC+"
	If basket_number = "X127FI2" then suggested_population = "LTC+"
	If basket_number = "X127FI7" then suggested_population = "LTC+"
	'Contacted Case Mgt
	If basket_number = "X127FG6" then suggested_population = "LTC+"           '"Kristen Kasem"
	If basket_number = "X127FG7" then suggested_population = "LTC+"           '"Kristen Kasem"
	If basket_number = "X127EM3" then suggested_population = "LTC+"           '"True L. or Gina G."
	If basket_number = "X127EM4" then suggested_population = "LTC+"            '"True L. or Gina G."
	If basket_number = "X127EW7" then suggested_population = "LTC+"            '"Kimberly Hill"
	If basket_number = "X127EW8" then suggested_population = "LTC+"            '"Kimberly Hill"
	If basket_number = "X127FF4" then suggested_population = "LTC+"            '"Alyssa Taylor"
	If basket_number = "X127FF5" then suggested_population = "LTC+"            '"Alyssa Taylor"

	If basket_number = "X127EH9" then suggested_population = "LTH"
	If basket_number = "X127EJ1" then suggested_population = "LTH"
	If basket_number = "X127EM2" then suggested_population = "LTH"
	If basket_number = "X127FE6" then suggested_population = "LTH"

	If basket_number = "X127FA5" then suggested_population = "YET"
	If basket_number = "X127FA6" then suggested_population = "YET"
	If basket_number = "X127FA7" then suggested_population = "YET"
	If basket_number = "X127FA8" then suggested_population = "YET"
	If basket_number = "X127FB1" then suggested_population = "YET"
	If basket_number = "X127FA9" then suggested_population = "YET"

	If trim(suggested_population) = "" then suggested_population = "No suggestions available"
end function
'THE SCRIPT-------------------------------------------------------------------------

EMConnect ""

'Checking for MAXIS
Call check_for_MAXIS(True)


const case_numb_col = 1
const update_date_col = 2
const curr_bskt_col = 3
const case_pop_col	= 4
const case_pnd2_col = 5
const FS_stat_col = 6
const MF_stat_col = 7
const DW_stat_col = 8
const GA_stat_col = 9
const MS_stat_col = 10
const GR_stat_col = 11
const CASH_PEND_stat_col = 12


year_info = DatePart("yyyy", date)
month_info = DatePart("m", date)
month_info = right("00" & month_info, 2)
day_info = DatePart("d", date)
day_info = right("00" & day_info, 2)
date_for_file = year_info & "-" & month_info & "-" & day_info

file_url = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Expedited SNAP - ES Workflow Information\Table - CasesPending - Unassigned Cases\" & date_for_file & " - CasesPending Unassigned.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
' Call navigate_to_MAXIS_screen("CCOL", "CLIC")

ObjExcel.Cells(1, curr_bskt_col).Value = "Current Basket"
ObjExcel.Cells(1, case_pop_col).Value = "Case Population"
ObjExcel.Cells(1,case_pnd2_col).Value = "Case PENDING"
ObjExcel.Cells(1, FS_stat_col).Value = "SNAP Status"
ObjExcel.Cells(1, MF_stat_col).Value ="MF Status"
ObjExcel.Cells(1, DW_stat_col).Value ="DW Status"
ObjExcel.Cells(1, GA_stat_col).Value ="GA Status"
ObjExcel.Cells(1, MS_stat_col).Value ="MS Status"
ObjExcel.Cells(1, GR_stat_col).Value ="GR Status"
ObjExcel.Cells(1, CASH_PEND_stat_col).Value ="Cash Pending"
ObjExcel.Rows(1).Font.Bold = True


MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, case_numb_col).Value)
	case_population = ""
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
	EMReadScreen curr_case_basket, 7, 21, 14
	Call find_q_flow_population(curr_case_basket, case_population)

	EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it will return true as privileged response.
	If priv_check <> "PRIVIL" THEN
		ObjExcel.Cells(excel_row, curr_bskt_col).Value = curr_case_basket
		ObjExcel.Cells(excel_row, case_pop_col).Value = case_population
		ObjExcel.Cells(excel_row, case_pnd2_col).Value  = case_pending
		ObjExcel.Cells(excel_row, FS_stat_col).Value = snap_status
		ObjExcel.Cells(excel_row, MF_stat_col).Value =mfip_status
		ObjExcel.Cells(excel_row, DW_stat_col).Value =dwp_status
		ObjExcel.Cells(excel_row, GA_stat_col).Value =ga_status
		ObjExcel.Cells(excel_row, MS_stat_col).Value =msa_status
		ObjExcel.Cells(excel_row, GR_stat_col).Value =grh_status
		ObjExcel.Cells(excel_row, CASH_PEND_stat_col).Value =unknown_cash_pending
	Else
		ObjExcel.Cells(excel_row, curr_bskt_col).Value = "PRIV"
	End If
	Call Back_to_SELF

	excel_row = excel_row + 1
Loop until trim(ObjExcel.Cells(excel_row, 1).Value) = ""

for i = 1 to 12
	ObjExcel.Columns(i).AutoFit()
Next

script_end_procedure("Thanks! We're done here.")
