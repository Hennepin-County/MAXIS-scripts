'Hard coding that needs to be updated each year: MAXIS_footer_year, counted_date_year 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - BANKED MONTHS REPORT.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 219         'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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

EMConnect ""		'connecting to MAXIS
Call get_county_code	'gets county name to input into the 1st col of the spreadsheet
developer_mode_checkbox = checked 	'defauting the person note option to NOT person note


'Runs the dialog'
Do
	Do
		Do
			'The dialog is defined in the loop as it can change as buttons are pressed (populating the dropdown)'
			BeginDialog dfln_selection_dialog, 0, 0, 266, 75, "Banked Month Report"
			  EditBox 15, 25, 190, 15, dfln_list_excel_file_path
			  ButtonGroup ButtonPressed
			    PushButton 215, 25, 45, 15, "Browse...", select_a_file_button
			  DropListBox 15, 55, 140, 15, "select one..." & sheet_list, worksheet_dropdown
			  ButtonGroup ButtonPressed
			    OkButton 175, 55, 40, 15
			    CancelButton 220, 55, 40, 15
			  Text 10, 10, 255, 10, "Select the Excel File that DHS provided with the list of Convicted Drug Felons."
			  Text 10, 45, 150, 10, "Select the correct worksheet in the Excel file:"
			EndDialog
			err_msg = ""
			Dialog dfln_selection_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then 
				If dfln_list_excel_file_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
					sheet_list = ""	'Blanks the Month list out so that the previous worksheets are not still included'
				End If 
				call file_selection_system_dialog(dfln_list_excel_file_path, ".xlsx") 'allows the user to select the file'
			End If 
			If dfln_list_excel_file_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(dfln_list_excel_file_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		sheet_list = ""
		For Each objWorkSheet In objWorkbook.Worksheets
			sheet_list = sheet_list & chr(9) & objWorkSheet.Name
		Next
		If worksheet_dropdown = "select one..." then err_msg = err_msg & vbNewLine & "You must select a month that you are running this script for."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

objExcel.worksheets(worksheet_dropdown).Activate			'Activates the selected worksheet'

excel_col = 1
Dim col_name_array
ReDim col_name_array(0)
Do
	ReDim Preserve col_name_array(excel_col - 1)
	col_name_array(excel_col - 1) = ucase(replace(objExcel.cells(1, excel_col).Value, " ", ""))
	excel_col = excel_col + 1
	end_of_list = objExcel.cells(1, excel_col).Value
Loop until end_of_list = ""

For each column in col_name_array
	MsgBox column
Next 
