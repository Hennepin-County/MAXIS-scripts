'STATS GATHERING=============================================================================================================
name_of_script = "BULK - Six Month MF ER List.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
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

excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Recertification Statistics\MF 6 Month ERs Possible List.xlsx"
call excel_open(excel_file_path, True, True, ObjExcel, ObjWorkbook)

Call check_for_MAXIS(True)

excel_row = 2
Do
	Call back_to_SELF

	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)

	Call navigate_to_MAXIS_screen("CASE", "CURR")
	row = 1
	col = 1
	EMSearch "MFIP", row, col
	EmReadscreen mfip_app_date, 8, row, 29
	ObjExcel.Cells(excel_row, 5).Value = mfip_app_date

	Call navigate_to_MAXIS_screen("STAT", "REVW")
	EMReadscreen last_revw_date, 8, 11, 37
	last_revw_date = replace(last_revw_date, "_", "")
	last_revw_date = trim(last_revw_date)
	last_revw_date = replace(last_revw_date, " ", "/")
	ObjExcel.Cells(excel_row, 6).Value = last_revw_date

	excel_row = excel_row + 1
	next_case_number = trim(ObjExcel.Cells(excel_row, 2).Value)
Loop until next_case_number = ""

ObjWorkbook.Save()
ObjExcel.ActiveWorkbook.Close
ObjExcel.Application.Quit
ObjExcel.Quit

Call script_end_procedure("ALL DONE")
