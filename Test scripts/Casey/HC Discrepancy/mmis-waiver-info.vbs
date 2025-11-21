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

function keep_MMIS_passworded_in(mmis_area, maxis_area)
    ' MsgBox running_stopwatch & vbNewLine & timer
    If timer - running_stopwatch > 720 Then         'this means the script has been running for more than 12 minutes since we last popped in to MMIS
        Call navigate_to_MMIS_region(mmis_area)      'Going to MMIS'
        'MsgBox "In MMIS"
        Call navigate_to_MAXIS(maxis_area)                       'going back to MAXIS'
        'MsgBox "Back to MAXIS"
        running_stopwatch = timer                                       'resetting the stopwatch'
    End If
end function

'THE SCRIPT-------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
running_stopwatch = timer               'setting the running timer so we log in to MMIS within every 15 mintues so we don't password out

file_selection_path = ""

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "Find Waiver information in MMIS"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 15, 20, 180, 20, "This script should be used on a list of cases that appear to have a waiver on HC in MAXIS."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
    Dialog Dialog1
    cancel_without_confirmation
    If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
Loop until ButtonPressed = OK and file_selection_path <> ""

' file_url = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\HC Discrepancy\LTC\Waiver Report.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_selection_path, visible_status, alerts_status, ObjExcel, objWorkbook)

mx_numb_col 				= "2"
pmi_numb_col 				= "4"
clt_name_col 				= "5"
MMIS_MA_Elig_End_col 		= "6"
MMIS_MA_Elig_Type_col		= "7"
MMIS_Waiver_col			    = "8"
MMIS_Waiver_End_Date_col	= "9"
MMIS_Waiver_Screen_col	    = "10"
end_in_future_col			= "11"


Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 225, "Find Waiver information in MMIS"
  Text 10, 15, 105, 10, "Is the PMI listed in the Excel?"
  DropListBox 115, 10, 135, 45, "Select One ..."+chr(9)+"Yes, PMI is known"+chr(9)+"No, PMI needs to be identified", get_pmi
  GroupBox 10, 30, 245, 140, "Column Identifiers"
  Text 15, 45, 135, 10, "Enter the column number for each type."
  Text 15, 55, 230, 10, "If any are missing, add them now and then enter the correct numbers."
  Text 15, 75, 75, 10, "Case Number Column"
  EditBox 90, 70, 25, 15, mx_numb_col
  Text 50, 95, 40, 10, "PMI Column"
  EditBox 90, 90, 25, 15, pmi_numb_col
  Text 45, 115, 45, 10, "Name Column"
  EditBox 90, 110, 25, 15, clt_name_col
  Text 20, 135, 70, 10, "MMIS ELIG End Date"
  EditBox 90, 130, 25, 15, MMIS_MA_Elig_End_col
  Text 35, 155, 55, 10, "MMIS ELIG Type"
  EditBox 90, 150, 25, 15, MMIS_MA_Elig_Type_col
  Text 150, 75, 65, 10, "MMIS Waiver Type "
  EditBox 215, 70, 25, 15, MMIS_Waiver_col
  Text 135, 95, 80, 10, "MMIS Waiver End Date"
  EditBox 215, 90, 25, 15, MMIS_Waiver_End_Date_col
  Text 125, 115, 90, 10, "MMIS Waiver Screen Date"
  EditBox 215, 110, 25, 15, MMIS_Waiver_Screen_col
  Text 150, 135, 65, 10, "Future End Column"
  EditBox 215, 130, 25, 15, end_in_future_col
  Text 10, 180, 225, 20, "Ensure you are logged into MMIS and if the PMI needs to be found also ensure you are logged into MAXIS."
  ButtonGroup ButtonPressed
    OkButton 150, 195, 50, 15
    CancelButton 205, 195, 50, 15
EndDialog

Do
    err_msg = ""
    Dialog Dialog1
    cancel_without_confirmation

    If get_pmi = "Select One ..." Then err_msg = err_msg & vbCr & "* Please select whether the PMI is listed in the Excel file or not."
    If NOT IsNumeric(mx_numb_col) Then err_msg = err_msg & vbCr & "* The Case Number Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(pmi_numb_col) Then err_msg = err_msg & vbCr & "* The PMI Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(clt_name_col) Then err_msg = err_msg & vbCr & "* The Client Name Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(MMIS_MA_Elig_End_col) Then err_msg = err_msg & vbCr & "* The MMIS MA Elig End Date Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(MMIS_MA_Elig_Type_col) Then err_msg = err_msg & vbCr & "* The MMIS MA Elig Type Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(MMIS_Waiver_col) Then err_msg = err_msg & vbCr & "* The MMIS Waiver Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(MMIS_Waiver_End_Date_col) Then err_msg = err_msg & vbCr & "* The MMIS Waiver End Date Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(MMIS_Waiver_Screen_col) Then err_msg = err_msg & vbCr & "* The MMIS Waiver Screen Date Column does not appear to be a number, all columns are required and should be entered as a number."
    If NOT IsNumeric(end_in_future_col) Then err_msg = err_msg & vbCr & "* The Future End Column does not appear to be a number, all columns are required and should be entered as a number."

    If err_msg <> "" Then
        MsgBox "Please correct the following errors:" & vbCr & err_msg, vbExclamation, "Error"
    End If
Loop until err_msg = ""

mx_numb_col = mx_numb_col * 1
pmi_numb_col = pmi_numb_col * 1
clt_name_col = clt_name_col * 1
MMIS_MA_Elig_End_col = MMIS_MA_Elig_End_col * 1
MMIS_MA_Elig_Type_col = MMIS_MA_Elig_Type_col * 1
MMIS_Waiver_col = MMIS_Waiver_col * 1
MMIS_Waiver_End_Date_col = MMIS_Waiver_End_Date_col * 1
MMIS_Waiver_Screen_col = MMIS_Waiver_Screen_col * 1
end_in_future_col = end_in_future_col * 1

If get_pmi = "No, PMI needs to be identified" Then
    Call check_for_MAXIS(False)
    Call back_to_SELF                                               'starting at the SELF panel
    EMReadScreen MX_environment, 13, 22, 48                         'seeing which MX environment we are in
    MX_environment = trim(MX_environment)
    Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")        'Going to MMIS'
    Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS

    adding_excel_row = 1 're-establishing the row to start checking the members for
    Do
        'Loops until there are no more cases in the Excel list
        adding_excel_row = adding_excel_row + 1
        MAXIS_case_number = objExcel.cells(adding_excel_row, mx_numb_col).Value   'reading the case number from Excel
        MAXIS_case_number = Trim(MAXIS_case_number)
		Call keep_MMIS_passworded_in("CTY ELIG STAFF/UPDATE", MX_environment)                'every 12 mintues or so, the script will pop in to MMIS to make sure we are passworded in
    Loop until MAXIS_case_number = ""
    final_row = adding_excel_row - 1

    For excel_row = 2 to final_row
        'Loops until there are no more cases in the Excel list
        MAXIS_case_number = objExcel.cells(excel_row, mx_numb_col).Value   'reading the case number from Excel
        MAXIS_case_number = Trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit for

        Call navigate_to_MAXIS_screen("CASE", "PERS")
        EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
        If PRIV_check <> "PRIV" then
            row = 10
            first_pers = True
            new_index = " "
            Do
                EMReadScreen HC_STAT, 1, row, 61
                If HC_STAT = "A" OR HC_STAT = "P" then
                    EMReadScreen person_PMI, 8, row, 34
                    EMReadscreen last_name, 15, row, 6
                    EMReadscreen first_name, 11, row, 22
                    If first_pers Then
                        objExcel.cells(excel_row, clt_name_col).Value = trim(first_name) & " " & trim(last_name)
                        objExcel.cells(excel_row, pmi_numb_col).Value = person_PMI

                        first_pers = False
                    Else
                        objExcel.cells(adding_excel_row, mx_numb_col).Value  = MAXIS_case_number
                        objExcel.cells(adding_excel_row, clt_name_col).Value = trim(first_name) & " " & trim(last_name)
                        objExcel.cells(adding_excel_row, pmi_numb_col).Value = person_PMI

                        adding_excel_row = adding_excel_row + 1
                    End If
                End If
                row = row + 3			'information is 3 rows apart. Will read for the next member.
                If row = 19 then
                    PF8
                    row = 10					'changes MAXIS row if more than one page exists
                END if
                EMReadScreen last_PERS_page, 21, 24, 2
            LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
        End If
        Call back_to_SELF
		Call keep_MMIS_passworded_in("CTY ELIG STAFF/UPDATE", MX_environment)                'every 12 mintues or so, the script will pop in to MMIS to make sure we are passworded in
    Next
    Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")
End If

'Checking for MMIS
Call check_for_MMIS(False)

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
excel_row = 2
Do
	MAXIS_case_number = right("00000000" & trim(ObjExcel.Cells(excel_row, mx_numb_col).Value), 8)
	PMI_number = right("00000000" & trim(ObjExcel.Cells(excel_row, pmi_numb_col).Value), 8)

    If PMI_number <> "00000000" Then
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
    Else
        ObjExcel.Cells(excel_row, MMIS_MA_Elig_End_col).Value = "NO PMI"
    End If

	excel_row = excel_row + 1
	next_case_numb = trim(ObjExcel.Cells(excel_row, mx_numb_col).Value)
Loop until next_case_numb = ""

Call script_end_procedure("DONE?")

