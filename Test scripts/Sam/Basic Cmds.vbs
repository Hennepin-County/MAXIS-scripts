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




Function file_selection_dialog()

'creates a Windows Script Host object
Set Fshell = CreateObject("WScript.Shell")

' creates a long string of powershell commands that will be executed
shellCmd = "powershell -NoProfile -NonInteractive -WindowStyle Hidden -command " & "Add-Type -AssemblyName System.Windows.Forms; " & _
            "$dlg = New-Object System.Windows.Forms.OpenFileDialog; " & _           
           "$dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop'); " & _
           "$dlg.Filter = 'Excel files (*.xlsx)|*.xlsx'; " & _ 
           "$dlg.ShowDialog() | Out-Null; " & _
           "$dlg.FileName; "


' Sets a variable of the file path selected from the PowerShell script run.
file_selection_path = Fshell.Exec(shellCmd).StdOut.ReadLine

end Function


'============= DIALOG BOX

EMConnect ""

Set file_selection_path = Nothing
Set selected_file = Nothing

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone and establishing county name
EMConnect ""

function write_date(date_variable, date_format_variable, screen_row, screen_col)
'--- This function will write a date in any format desired.
'~~~~~ date_variable: date to write
'~~~~~ date_format_variable: format of date. this should be a string with the correct spaces between month/day/year examples: MM DD YY, MM YY, MM  DD  YYYY
'~~~~~ screen_row: row to write date
'~~~~~ screen_col: column to write date
'===== Keywords: MAXIS, MMIS, PRISM, date, format
    'Figures out the format of the month. If it was "MM", "M", or not present.
    If instr(ucase(date_format_variable), "MM") <> 0 then
        month_format = "MM"
        month_position = instr(ucase(date_format_variable), "MM")
    Elseif instr(ucase(date_format_variable), "M") <> 0 then
        month_format = "M"
        month_position = instr(ucase(date_format_variable), "M")
    Else
        month_format = ""
        month_position = 0
    End if

    'Figures out the format of the day. If it was "DD", "D", or not present.
    If instr(ucase(date_format_variable), "DD") <> 0 then
        day_format = "DD"
        day_position = instr(ucase(date_format_variable), "DD")
    Elseif instr(ucase(date_format_variable), "D") <> 0 then
        day_format = "D"
        day_position = instr(ucase(date_format_variable), "D")
    Else
        day_format = ""
        day_position = 0
    End if

    'Figures out the format of the year. If it was "YYYY", "YY", or not present.
    If instr(ucase(date_format_variable), "YYYY") <> 0 then
        year_format = "YYYY"
        year_position = instr(ucase(date_format_variable), "YYYY")
    Elseif instr(ucase(date_format_variable), "YY") <> 0 then
        year_format = "YY"
        year_position = instr(ucase(date_format_variable), "YY")
    Else
        year_format = ""
        year_position = 0
    End if

    'Formats the month. Separates the month into its own variable and adds a zero if needed.
    var_month = datepart("m", date_variable)
    IF len(var_month) = 1 and month_format = "MM" THEN var_month = "0" & var_month

    'Formats the day. Separates the day into its own variable and adds a zero if needed.
    var_day = datepart("d", date_variable)
    IF len(var_day) = 1 and day_format = "DD" THEN var_day = "0" & var_day

    'Formats the year based on "YY" or "YYYY" formatting.
    If year_format = "YY" then
        var_year = right(datepart("yyyy", date_variable), 2)
    ElseIf year_format = "YYYY" then
        var_year = datepart("yyyy", date_variable)
    END IF

    If month_position <> 0 Then EMWriteScreen var_month, screen_row, screen_col + month_position - 1
    If day_position <> 0 Then EMWriteScreen var_day, screen_row, screen_col + day_position - 1
    If year_position <> 0 Then EMWriteScreen var_year, screen_row, screen_col + year_position - 1
end function

'assigning current month and year for MAXIS navigation footer months; establishing current day for referral date
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
CM_day = right("0" &             DatePart("d",           date                             ), 2)
date_of_today = CM_mo & "/" & CM_day & "/" & CM_yr


msgBox(date_of_today)

		'call navigate_to_MAXIS_screen_review_PRIV("pers", clear_line_of_text(18, 43), is_this_priv)
			'If is_this_priv = false then msgbox "hello world" : call script_end_procedure("you did two things at once!") 





Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "CBO referral"
    ButtonGroup ButtonPressed
    PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 145, 90, 50, 15
    CancelButton 200, 90, 50, 15
    EditBox 15, 45, 180, 15, file_selection_path
    GroupBox 10, 5, 250, 80, "Using the SEND MANUAL REFERRAL script"
    Text 20, 20, 235, 20, "This script should be used when E & T provides you with a list of recipeints that are working with CBO's and a manual referral is needed. "
    Text 15, 65, 230, 15, "Select the Excel file that contains the CBO information by selecting the 'Browse' button, and finding the file."
EndDialog

'dialog and dialog DO...Loop
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_dialog()
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    

    
