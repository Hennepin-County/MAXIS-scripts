'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - EXPEDITED DETERMINATION REPORT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	60			 'manual run time in seconds
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
call changelog_update("10/15/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

exp_assignment_folder = t_drive & "\Eligibility Support\Assignments\Expedited Information"
Set objFolder = objFSO.GetFolder(exp_assignment_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

'Open an existing Excel for the year
report_out_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\2021 EXP Determination Report Out.xlsx"

Call excel_open(report_out_file, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

For Each objWorkSheet In objWorkbook.Worksheets
    If instr(objWorkSheet.Name, "ALL CASES") <> 0 Then
        objWorkSheet.Activate
        Exit For
    End If
Next
total_excel_row = 1
Do
    total_excel_row = total_excel_row + 1
    this_case_number = trim(ObjExcel.Cells(total_excel_row, 1).Value)
Loop Until this_case_number = ""

For Each objFile in colFiles																'looping through each file

    this_file_path = objFile.Path
    ' MsgBox this_file_path
    'Setting the object to open the text file for reading the data already in the file
    Set objTextStream = objFSO.OpenTextFile(this_file_path, ForReading)

    'Reading the entire text file into a string
    every_line_in_text_file = objTextStream.ReadAll

    exp_det_details = split(every_line_in_text_file, vbNewLine)

    For Each text_line in exp_det_details
        If Instr(text_line, "^*^*^") <> 0 Then
            line_info = split(text_line, "^*^*^")
            line_info(0) = trim(line_info(0))
            If line_info(0) = "CASE NUMBER" Then ObjExcel.Cells(total_excel_row, 1).Value = line_info(1)
            If line_info(0) = "WORKER NAME" Then ObjExcel.Cells(total_excel_row, 2).Value = line_info(1)
            If line_info(0) = "CASE X NUMBER" Then ObjExcel.Cells(total_excel_row, 3).Value = line_info(1)
            If line_info(0) = "DATE OF APPLICATION" Then ObjExcel.Cells(total_excel_row, 4).Value = line_info(1)
            If line_info(0) = "DATE OF INTERVIEW" Then ObjExcel.Cells(total_excel_row, 5).Value = line_info(1)
            If line_info(0) = "EXPEDITED SCREENING STATUS" Then ObjExcel.Cells(total_excel_row, 6).Value = line_info(1)
            If line_info(0) = "EXPEDITED DETERMINATION STATUS" Then ObjExcel.Cells(total_excel_row, 7).Value = line_info(1)
            If line_info(0) = "DATE OF APPROVAL" Then ObjExcel.Cells(total_excel_row, 8).Value = line_info(1)
            If line_info(0) = "SNAP DENIAL DATE" Then ObjExcel.Cells(total_excel_row, 9).Value = line_info(1)
            If line_info(0) = "SNAP DENIAL REASON" Then ObjExcel.Cells(total_excel_row, 10).Value = line_info(1)
            If line_info(0) = "ID ON FILE" Then ObjExcel.Cells(total_excel_row, 11).Value = line_info(1)
            If line_info(0) = "END DATE OF SNAP IN ANOTHER STATE" Then ObjExcel.Cells(total_excel_row, 12).Value = line_info(1)
            If line_info(0) = "EXPEDITED APPROVE PREVIOUSLY POSTPONED" Then ObjExcel.Cells(total_excel_row, 13).Value = line_info(1)
            If line_info(0) = "EXPLAIN APPROVAL DELAYS" Then ObjExcel.Cells(total_excel_row, 14).Value = line_info(1)
            If line_info(0) = "POSTPONED VERIFICATIONS" Then ObjExcel.Cells(total_excel_row, 15).Value = line_info(1)
            If line_info(0) = "WHAT ARE THE POSTPONED VERIFICATIONS" Then ObjExcel.Cells(total_excel_row, 16).Value = line_info(1)
            If line_info(0) = "DATE OF SCRIPT RUN" Then ObjExcel.Cells(total_excel_row, 17).Value = line_info(1)
        End If
    Next
    total_excel_row = total_excel_row + 1
Next

'Add a sheet to the Excel with the report date
sheet_friendly_date = replace(date, "/", "-")
sheet_name = sheet_friendly_date & " REPT"
ObjExcel.Worksheets.Add().Name = sheet_name

'ADD HEADERS HERE'
ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
ObjExcel.Cells(1, 2).Value = "WORKER NAME"
ObjExcel.Cells(1, 3).Value = "CASE X NUMBER"
ObjExcel.Cells(1, 4).Value = "DATE OF APPLICATION"
ObjExcel.Cells(1, 5).Value = "DATE OF INTERVIEW"
ObjExcel.Cells(1, 6).Value = "EXPEDITED SCREENING STATUS"
ObjExcel.Cells(1, 7).Value = "EXPEDITED DETERMINATION STATUS"
ObjExcel.Cells(1, 8).Value = "DATE OF APPROVAL"
ObjExcel.Cells(1, 9).Value = "SNAP DENIAL DATE"
ObjExcel.Cells(1, 10).Value = "SNAP DENIAL REASON"
ObjExcel.Cells(1, 11).Value = "ID ON FILE"
ObjExcel.Cells(1, 12).Value = "END DATE OF SNAP IN ANOTHER STATE"
ObjExcel.Cells(1, 13).Value = "EXPEDITED APPROVE PREVIOUSLY POSTPONED" 				'(Boolean)
ObjExcel.Cells(1, 14).Value = "EXPLAIN APPROVAL DELAYS " 								'(all of them)
ObjExcel.Cells(1, 15).Value = "POSTPONED VERIFICATIONS"
ObjExcel.Cells(1, 16).Value = "WHAT ARE THE POSTPONED VERIFICATIONS"
ObjExcel.Cells(1, 17).Value = "DATE OF SCRIPT RUN"
ObjExcel.Rows(1).Font.Bold = True

excel_row = 2
'Create an array of all of the files in the folder

For Each objFile in colFiles																'looping through each file

    this_file_path = objFile.Path
    ' MsgBox this_file_path
    'Setting the object to open the text file for reading the data already in the file
    Set objTextStream = objFSO.OpenTextFile(this_file_path, ForReading)

    'Reading the entire text file into a string
    every_line_in_text_file = objTextStream.ReadAll

    exp_det_details = split(every_line_in_text_file, vbNewLine)

    For Each text_line in exp_det_details
        If Instr(text_line, "^*^*^") <> 0 Then
            line_info = split(text_line, "^*^*^")
            line_info(0) = trim(line_info(0))
            If line_info(0) = "CASE NUMBER" Then ObjExcel.Cells(excel_row, 1).Value = line_info(1)
            If line_info(0) = "WORKER NAME" Then ObjExcel.Cells(excel_row, 2).Value = line_info(1)
            If line_info(0) = "CASE X NUMBER" Then ObjExcel.Cells(excel_row, 3).Value = line_info(1)
            If line_info(0) = "DATE OF APPLICATION" Then ObjExcel.Cells(excel_row, 4).Value = line_info(1)
            If line_info(0) = "DATE OF INTERVIEW" Then ObjExcel.Cells(excel_row, 5).Value = line_info(1)
            If line_info(0) = "EXPEDITED SCREENING STATUS" Then ObjExcel.Cells(excel_row, 6).Value = line_info(1)
            If line_info(0) = "EXPEDITED DETERMINATION STATUS" Then ObjExcel.Cells(excel_row, 7).Value = line_info(1)
            If line_info(0) = "DATE OF APPROVAL" Then ObjExcel.Cells(excel_row, 8).Value = line_info(1)
            If line_info(0) = "SNAP DENIAL DATE" Then ObjExcel.Cells(excel_row, 9).Value = line_info(1)
            If line_info(0) = "SNAP DENIAL REASON" Then ObjExcel.Cells(excel_row, 10).Value = line_info(1)
            If line_info(0) = "ID ON FILE" Then ObjExcel.Cells(excel_row, 11).Value = line_info(1)
            If line_info(0) = "END DATE OF SNAP IN ANOTHER STATE" Then ObjExcel.Cells(excel_row, 12).Value = line_info(1)
            If line_info(0) = "EXPEDITED APPROVE PREVIOUSLY POSTPONED" Then ObjExcel.Cells(excel_row, 13).Value = line_info(1)
            If line_info(0) = "EXPLAIN APPROVAL DELAYS" Then ObjExcel.Cells(excel_row, 14).Value = line_info(1)
            If line_info(0) = "POSTPONED VERIFICATIONS" Then ObjExcel.Cells(excel_row, 15).Value = line_info(1)
            If line_info(0) = "WHAT ARE THE POSTPONED VERIFICATIONS" Then ObjExcel.Cells(excel_row, 16).Value = line_info(1)
            If line_info(0) = "DATE OF SCRIPT RUN" Then ObjExcel.Cells(excel_row, 17).Value = line_info(1)
        End If
    Next
    excel_row = excel_row + 1

    ' objFSO.DeleteFile(this_file_path)

Next

Const xlSrcRange = 1
Const xlYes = 1
xlVAlignTop = -4160
xlHAlignLeft = -4131
For col = 1 to 17
    ObjExcel.columns(col).AutoFit()
    ObjExcel.columns(col).VerticalAlignment = xlVAlignTop
    ObjExcel.columns(col).HorizontalAlignment = xlHAlignLeft
Next


ObjExcel.Columns(10).ColumnWidth = 150
ObjExcel.Columns(10).WrapText = True
ObjExcel.Columns(14).ColumnWidth = 150
ObjExcel.Columns(14).WrapText = True

tableRange = "A1:Q" & excel_row-1
table_friendly_date = replace(date, "/", "")
table_friendly_date = trim(table_friendly_date)
table_name = table_friendly_date & "TABLE"
ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, tableRange, xlYes).Name = table_name
' ObjExcel.ActiveSheet.ListObjects(table_name).TableStyle = "TableStyleDark2"

'Loop through each one
    'read the files one by one
    'Add detail of the files to the Excel sheet
'Update statistics in the Excel

For Each objWorkSheet In objWorkbook.Worksheets
    If instr(objWorkSheet.Name, "Statistics") <> 0 Then
        objWorkSheet.Activate
        Exit For
    End If
Next



'SAVE EXCEL'

MsgBox "DONE"
