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

curr_year = "2025"
curr_month = "April"
manager_log_file_path 				= "https://hennepin.sharepoint.com/teams/InterviewPhoneHSRs-Supervisors/Shared%20Documents/Management%20Team-HSR%20Interviews/Manager%20Log%20" & curr_year & "%20" & curr_month & ".xlsx"
mnger_log_processor_col = 5
start_excel = 1596
end_excel = 1619

Call excel_open(manager_log_file_path, True, False, ObjMngrExcel, objMngrWorkbook)
ObjMngrExcel.worksheets("Manager log").Activate
'(1598-1616)


For excel_row = start_excel to end_excel
	If trim(ObjMngrExcel.Cells(excel_row, mnger_log_processor_col).Value) = "" Then
		ObjMngrExcel.Cells(excel_row, mnger_log_processor_col).Value = "=IFERROR(XLOOKUP(1, (T_Processors[Area  (SORT-4)]=[@Area])*(T_Processors[Match HSR '#]=[@[Match HSR]]), T_Processors[Processors  (SORT-1)]),  " & Chr(34) & Chr(34) & ")"
	End If
Next

MsgBox "WAIT HERE"
stopscript