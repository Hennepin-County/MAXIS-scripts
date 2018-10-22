'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "bulk-applications.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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

' 'Reading Locally held FuncLib in leiu of issues with connecting to GitHub
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs")
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script
'----------------------------------------------------------------------------------------------------Custom function
function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
end function

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("02/05/2018", "Initial version.", "MiKayla Handley, Hennepin County")


'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

'dialog and dialog DO...Loop
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
		BeginDialog file_select_dialog, 0, 0, 221, 50, "Select the source file"
  			ButtonGroup ButtonPressed
    		PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    		OkButton 110, 30, 50, 15
    		CancelButton 165, 30, 50, 15
  			EditBox 5, 10, 165, 15, file_selection_path
		EndDialog
		err_msg = ""
		Dialog file_select_dialog
		If ButtonPressed = cancel then stopscript
		If ButtonPressed = select_a_file_button then
			If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


entry_row = 4
stats_header_col = 17
stats_col = 18
date_to_assess = #03/21/18#
thirty_days_ago = DateAdd("d", -30, date_to_assess)


objExcel.Cells(entry_row, stats_header_col).Value       = "Cases at 30 DAYS"        'number of notices that were successful
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(G:G," & Chr(34)  & thirty_days_ago & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Cases OVER 30 DAYS"        'number of notices that were successful
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(G:G," & Chr(34) & "<" & thirty_days_ago & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Cases at potential Denial"           'calculation of the percent of successful notices
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(L:L," & Chr(34) & "TRUE" & Chr(34) & ")"
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Appointment Notices Sent"           'calculation of the percent of successful notices
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(I:I," & Chr(34) & date_to_assess & Chr(34) & ",K:K," & Chr(34) & "Y" & Chr(34) & ")"
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "NOMIs Sent"           'calculation of the percent of successful notices
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(J:J," & Chr(34) & date_to_assess & Chr(34) & ",K:K," & Chr(34) & "Y" & Chr(34) & ")"

entry_row = entry_row + 1

for row_to_change = 1 to entry_row
    objExcel.Cells(row_to_change, stats_header_col).font.colorindex = 1
    objExcel.Cells(row_to_change, stats_col).font.colorindex = 1
next
