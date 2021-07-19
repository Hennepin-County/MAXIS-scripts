'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - FORM FROM DHS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 150                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
Call changelog_update("11/20/2020", "Added stars to the NOTE for a delimitator between notes.", "Casey Love, Hennepin County")
call changelog_update("11/10/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

Call check_for_MMIS(True)

'call check_for_MMIS(True) 'Sending MMIS back to the beginning screen and checking for a password prompt
Call MMIS_case_number_finder(MMIS_case_number)

Call get_to_RKEY

BeginDialog Dialog1, 0, 0, 256, 95, "Dialog"
  EditBox 75, 10, 70, 15, MMIS_case_number
  DropListBox 75, 30, 85, 45, "MSHO", form_type
  EditBox 75, 50, 175, 15, worker_signature
  DropListBox 75, 70, 60, 45, "Select"+chr(9)+"Yes"+chr(9)+"No", faxed_yn
  ButtonGroup ButtonPressed
    OkButton 145, 75, 50, 15
    CancelButton 200, 75, 50, 15
  Text 20, 15, 50, 10, "Case Number:"
  Text 30, 35, 40, 10, "Form Type:"
  Text 10, 55, 60, 10, "Worker Signature:"
  Text 15, 75, 50, 10, "Faxed to DHS?"
EndDialog
'do the dialog here
Do
    err_msg = ""

	Dialog Dialog1
	cancel_without_confirmation

    MMIS_case_number = trim(MMIS_case_number)
	worker_signature = trim(worker_signature)

	If MMIS_case_number = "" Then err_msg = err_msg & vbNewLine & "* Enter a case number to run this script on."
	If worker_signature = "" Then err_msg = err_msg & vbNewLine & "* Sign your case note."
	If faxed_yn = "Select" Then err_msg = err_msg & vbNewLine & "* Indicate if a fax is being sent to DHS."

    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

'checking for an active MMIS session
Call check_for_MMIS(True)
Call get_to_RKEY

MMIS_case_number = right("00000000" & MMIS_case_number, 8)

'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
EMWriteScreen "c", 2, 19
Call clear_line_of_text(4, 19)		'Clearing all of the search options used on RKEY as we must ONLY enter a case number
Call clear_line_of_text(5, 19)
Call clear_line_of_text(5, 48)
Call clear_line_of_text(6, 19)
Call clear_line_of_text(6, 48)
Call clear_line_of_text(6, 69)
Call clear_line_of_text(9, 19)
Call clear_line_of_text(9, 48)
Call clear_line_of_text(9, 69)

EMWriteScreen MMIS_case_number, 9, 19
transmit
transmit
transmit

EMReadScreen clt_last_name, 17, 11, 24
EMReadScreen clt_first_name, 10, 11, 42
clt_last_name = trim(clt_last_name)
clt_first_name = trim(clt_first_name)
client_name = clt_first_name & " " & clt_last_name

transmit

pf4
pf11		'Starts a new case note'

If form_type = "MSHO" Then
	If faxed_yn = "Yes" Then
		CALL write_variable_in_MMIS_NOTE("MSHO AHPS enrollment form received by Hennepin County for " & client_name & ".")
		CALL write_variable_in_MMIS_NOTE("Faxed to DHS-MSHO")
	ElseIf faxed_yn = "No" Then
		CALL write_variable_in_MMIS_NOTE("AHPS MSHO form received by Hennepin with no plan change requested.")
		CALL write_variable_in_MMIS_NOTE("No action taken.")
	End If
End If
CALL write_variable_in_MMIS_NOTE(worker_signature)
CALL write_variable_in_MMIS_NOTE ("*************************************************************************")

pf3
pf3
IF REFM_error_check = "WARNING: MA12,01/16" Then
	PF3
END IF

If faxed_yn = "Yes" Then
	mhc_msho_list = t_drive & "\Eligibility Support\EA_ADAD\EA_ADAD_MHC\Forms Faxed List.xlsx"

	Call find_user_name(worker_name)						'defaulting the name of the suer running the script
	If worker_name = "Casey H Love" Then mhc_msho_list = "C:\MAXIS-scripts\Forms Faxed List.xlsx"
	'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
	call excel_open(mhc_msho_list, FALSE, True, ObjExcel, objWorkbook)

	excel_row = 1
	Do
		excel_row = excel_row + 1
		row_information = trim(ObjExcel.Cells(excel_row, 1).Value)
	Loop until row_information = ""

	ObjExcel.Cells(excel_row, 1).Value = MMIS_case_number
	ObjExcel.Cells(excel_row, 2).Value = form_type
	ObjExcel.Cells(excel_row, 3).Value = date
	ObjExcel.Cells(excel_row, 4).Value = time
	ObjExcel.Cells(excel_row, 5).Value = worker_name

	ObjExcel.ActiveWorkbook.Save                                            'saving and closing the Excel spreadsheet
	ObjExcel.ActiveWorkbook.Close
	ObjExcel.Application.Quit
End If
MAXIS_case_number = MMIS_case_number

call script_end_procedure_with_error_report("Success!! NOTE entered into MMIS regarding faxing the form to DHS.")
