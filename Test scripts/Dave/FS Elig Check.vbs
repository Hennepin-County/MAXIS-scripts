
'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - AVS Panel Report.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")

call changelog_update("04/25/2025", "Initial version.", "Dave Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Constants

case_col   = 2
date_col   = 5
status_col = 7
prog_col   = 4
version_date_col = 8
maxis_footer_month = "07"
maxis_footer_year = "25"
'FILE SELECTION DIALOG
DO
	call file_selection_system_dialog(excel_file_path, ".xlsx")
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(excel_file_path)
	objExcel.Visible = True
	objExcel.DisplayAlerts = True
	confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
	IF confirm_file = vbCancel THEN
		objWorkbook.Close
		objExcel.Quit
		stopscript
	ELSEIF confirm_file = vbNo THEN
		objWorkbook.Close
		objExcel.Quit
	END IF
LOOP UNTIL confirm_file = vbYes

Call check_for_MAXIS(true)






'For each row in the sheet
excel_row = 42261 'start at row 2, row 1 is the header

Do while ObjExcel.Cells(excel_row, case_col).value <> ""
    MAXIS_case_number = ObjExcel.Cells(excel_row, case_col).value
	progr = ObjExcel.Cells(excel_row, prog_col).value
	progr = trim(progr)
	'msgbox ObjExcel.Cells(excel_row, prog_col).value
	'msgbox progr
	If progr <> "PX" Then
		If progr = "FS" OR progr = "EA" OR progr = "RC" Then
			prog_row = 19
			prog_col = 78
		ElseIf progr = "4E" OR progr = "DW" OR progr = "EG" OR progr = "GR" OR progr = "MF" OR progr = "MS" Then
			prog_row = 20
			prog_col = 79
		End If
		IF progr = "FS" Then prog_elig = "FS"
		IF progr = "GA" Then prog_elig = "GA"
		IF progr = "MF" Then prog_elig = "MF"
		If progr = "EA" or progr = "EG" Then prog_elig = "EMER"
		If progr = "4E" Then prog_elig = "IVE"
		If progr = "GR" Then prog_elig = "GRH"
		If progr = "MS" Then prog_elig = "MSA"
		If progr = "RC" Then prog_elig = "RCA"
		If progr = "DW" THEN prog_elig = "DWP"

    	Call convert_date_into_MAXIS_footer_month(ObjExcel.Cells(excel_row, date_col).value, maxis_footer_month, maxis_footer_year)
    	Call navigate_to_MAXIS_screen_review_priv("ELIG", prog_elig, is_this_priv)
    	If is_this_priv = false Then
			EmReadScreen error_check, 5, 24, 2
			If trim(error_check) = "" Then
    	    	call find_last_approved_ELIG_version(prog_row, prog_col, version_number, version_date, version_result, approval_found)

    	    	If approval_found = true Then
    	    	    ObjExcel.Cells(excel_row, status_col) = version_result
    	    	    ObjExcel.Cells(excel_row, version_date_col) = version_date
    	    	Else
    	    	    ObjExcel.Cells(excel_row, status_col) = "No App Found"
    	    	End If
			Else
			ObjExcel.Cells(excel_row, status_col) = "Error"
			End If
    	Else
    	    objExcel.Cells(excel_row, status_col) = "Priv Case"
    	End If
	End If

    excel_row = excel_row + 1
Loop
'If the case number is not blank, then
'   go to the case number
'   Read the AREP name / address
'   If the AREP name is not blank, then
'      Add to sheet
   ' End If
'End If
   'Next

script_end_procedure()
