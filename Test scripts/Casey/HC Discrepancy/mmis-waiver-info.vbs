'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "READ MMIS WAIVER INFO FROM LIST.vbs"
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


'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Connects to BlueZone
EMConnect ""

'Checking for MMIS
Call check_for_MMIS(True)

file_url = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\HC Discrepancy\LTC\Waiver Report.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)


const mx_numb_col 				= 2
const pmi_numb_col 				= 5
const MMIS_MA_Elig_End_col 		= 9
const MMIS_MA_Elig_Type_col		= 10
const MMIS_Waiver_col			= 11
const MMIS_Waiver_End_Date_col	= 12
const MMIS_Waiver_Screen_col	= 13
const end_in_future_col			= 14


MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
excel_row = 2
Do
	MAXIS_case_number = right("00000000" & trim(ObjExcel.Cells(excel_row, mx_numb_col).Value), 8)
	PMI_number = right("00000000" & trim(ObjExcel.Cells(excel_row, pmi_numb_col).Value), 8)

	Call get_to_RKEY

	'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
	EMWriteScreen "I", 2, 19
	EMWriteScreen "        ", 9, 19
	Call write_value_and_transmit(PMI_number, 4, 19)
	EMReadscreen RKEY_check, 4, 1, 52
	If RKEY_check <> "RKEY" then
		EMReadScreen maj_prog, 2, 6, 13
		EMReadScreen maj_type, 2, 6, 35
		EMReadScreen maj_end, 8, 7, 54
		EMReadScreen waiv_type, 1, 15, 15
		EMReadScreen waiv_end, 8, 15, 46
		EMReadScreen waiv_screen, 8, 15, 71

		ObjExcel.Cells(excel_row, MMIS_MA_Elig_End_col).Value = maj_end
		ObjExcel.Cells(excel_row, MMIS_MA_Elig_Type_col).Value = maj_prog & " - " & maj_type
		ObjExcel.Cells(excel_row, MMIS_Waiver_col).Value = waiv_type
		ObjExcel.Cells(excel_row, MMIS_Waiver_End_Date_col).Value = waiv_end
		ObjExcel.Cells(excel_row, MMIS_Waiver_Screen_col).Value = waiv_screen

		If IsDate(waiv_end) Then

			If DateDiff("d", date, waiv_end) > 0 Then
				ObjExcel.Cells(excel_row, end_in_future_col).Value = "True"
			Else
				ObjExcel.Cells(excel_row, end_in_future_col).Value = "False"
			End If

		Else
			if waiv_end = "99/99/99" Then
				ObjExcel.Cells(excel_row, end_in_future_col).Value = "True"
			Else
				ObjExcel.Cells(excel_row, end_in_future_col).Value = "False"
			End If
		End If
	Else
		ObjExcel.Cells(excel_row, MMIS_MA_Elig_End_col).Value = "FAIL"
	End If

	excel_row = excel_row + 1
	next_case_numb = trim(ObjExcel.Cells(excel_row, mx_numb_col).Value)
Loop until next_case_numb = ""

Call script_end_procedure("DONE?")

