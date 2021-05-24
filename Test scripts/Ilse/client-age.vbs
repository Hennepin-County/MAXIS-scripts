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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'Function client_age(client_DOB, client_months_and_days)
'    Dim CurrentDate, Years, ThisYear, Months, ThisMonth, Days
'    CurrentDate = cdate(client_DOB)
'    Years = DateDiff("yyyy", CurrentDate, Date)
'    ThisYear = DateAdd("yyyy", Years, CurrentDate)
'    Months = DateDiff("m", ThisYear, Date)
'    ThisMonth = DateAdd("m", Months, ThisYear)
'    Days = DateDiff("d", ThisMonth, Date)
'
'    Do While (Days < 0) Or (Months < 0)
'        If Days < 0 Then
'            Months = Months - 1
'            ThisMonth = DateAdd("m", Months, ThisYear)
'            Days = DateDiff("d", ThisMonth, Date)
'        End If
'        If Months < 0 Then
'            Years = Years - 1
'            ThisYear = DateAdd("yyyy", Years, CurrentDate)
'            Months = DateDiff("m", ThisYear, Date)
'            ThisMonth = DateAdd("m", Months, ThisYear)
'            Days = DateDiff("d", ThisMonth, Date)
'        End If
'    Loop
'
'    age_of_client = Years
'    If client_months_and_days = True then age_of_client = age_of_client & "y/" & Months & "m/" & Days & "/d"
'End Function

file_selection_path = "C:\Users\ilfe001\OneDrive - Hennepin County\Desktop\Active GA Recipients.xlsx"

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "BULK - ABAWD REPORT"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of SNAP cases wtih member numbers are provided by BOBI to gather ABAWD, FSET and Banked Months information."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

do
    dialog Dialog1
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart
update_count = 0

Do
	client_DOB = ObjExcel.Cells(excel_row, 5).Value
	client_DOB = trim(client_DOB)

    MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	MAXIS_case_number = trim(MAXIS_case_number)

	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
    If snap_case = True then
        ObjExcel.Cells(excel_row, 9).Value = True
    Else
        ObjExcel.Cells(excel_row, 9).Value = False
    End if

    'Dim CurrentDate, Years, ThisYear, Months, ThisMonth, Days
    CurrentDate = CDate(client_DOB)
    Years = DateDiff("yyyy", CurrentDate, Date)
    ThisYear = DateAdd("yyyy", Years, CurrentDate)
    Months = DateDiff("m", ThisYear, Date)
    ThisMonth = DateAdd("m", Months, ThisYear)
    Days = DateDiff("d", ThisMonth, Date)

    Do While (Days < 0) Or (Months < 0)
        If Days < 0 Then
            Months = Months - 1
            ThisMonth = DateAdd("m", Months, ThisYear)
            Days = DateDiff("d", ThisMonth, Date)
        End If
        If Months < 0 Then
            Years = Years - 1
            ThisYear = DateAdd("yyyy", Years, CurrentDate)
            Months = DateDiff("m", ThisYear, Date)
            ThisMonth = DateAdd("m", Months, ThisYear)
            Days = DateDiff("d", ThisMonth, Date)
        End If
    Loop
    client_age = Years ''& "y/" & Months & "m/" & Days
    ObjExcel.Cells(excel_row, 6).Value = client_age

    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 5).Value = ""

msgbox "all done"
stopscript
