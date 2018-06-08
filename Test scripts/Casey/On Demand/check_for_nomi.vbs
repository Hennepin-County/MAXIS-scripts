'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - Check for NOMI.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 304			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'DIALOGS ===================================================================================================================

'Initial Dialog which requests a file path for the excel file
BeginDialog recert_list_dlg, 0, 0, 361, 105, "On Demand Recertifications"
  EditBox 130, 60, 175, 15, recertification_cases_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 310, 60, 45, 15, "Browse...", select_a_file_button
  EditBox 75, 85, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 85, 50, 15
    CancelButton 305, 85, 50, 15
  Text 10, 10, 170, 10, "Welcome to the On Demand Recertification Notifier."
  Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
  Text 10, 65, 120, 10, "Select an Excel file for recert cases:"
  Text 10, 90, 60, 10, "Worker Signature"
EndDialog

Do
	Dialog recert_list_dlg
	If ButtonPressed = cancel then stopscript
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(recertification_cases_excel_file_path, ".xlsx")
Loop until ButtonPressed = OK and recertification_cases_excel_file_path <> "" and worker_signature <> ""

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(recertification_cases_excel_file_path, True, True, ObjExcel, objWorkbook)
'Activates worksheet based on user selection
objExcel.worksheets("All cases").Activate

MAXIS_footer_month = "04"
MAXIS_footer_year = "18"

excel_row = 2
case_number_col = 2
nomi_success_col = 10
interview_col = 12

'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
today_mo = DatePart("m", date)
today_mo = right("00" & today_mo, 2)

today_day = DatePart("d", date)
today_day = right("00" & today_day, 2)

today_yr = DatePart("yyyy", date)
today_yr = right(today_yr, 2)

today_date = today_mo & "/" & today_day & "/" & today_yr

Do
    case_number = objExcel.Cells(excel_row, case_number_col).Value
    case_number = trim(case_number)
    nomi_info = objExcel.Cells(excel_row, nomi_success_col).Value
    nomi_info = trim(nomi_info)
    interview_date =  objExcel.Cells(excel_row, interview_col).Value
    interview_date = trim(interview_date)

    If nomi_info = "" Then
        If case_number <> "" and interview_date = "" Then
            MAXIS_case_number = case_number

            Call navigate_to_MAXIS_screen ("SPEC", "MEMO")


            memo_row = 7                                            'Setting the row for the loop to read MEMOs
            notc_confirm = "N"         'Defaulting this to 'N'
            Do
                EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
                EMReadScreen print_status, 7, memo_row, 67
                If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
                    notc_confirm = "Y"             'If we've found this then no reason to keep looking.
                    Exit Do
                End If

                memo_row = memo_row + 1           'Looking at next row'
            Loop Until create_date = "        "

            objExcel.Cells(excel_row, nomi_success_col).Value = notc_confirm

        End If
    End If

    excel_row = excel_row + 1
    'MsgBox excel_row
    back_to_self
Loop until case_number = ""
