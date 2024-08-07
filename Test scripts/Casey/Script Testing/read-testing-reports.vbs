'Required for statistical purposes==========================================================================================
name_of_script = "CL TEST - Script Test Log Report.vbs"
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


const script_run_date_time_col	= 1
const script_user_col			= 2
const file_name_col				= 3
const confrim_numb_col			= 4
const appl_date_and_time		= 5
const script_run_time			= 6
const panel_title_end			= 7
const search_type_col			= 8
const search_found_col			= 9
const search_notes_col			= 10
const xml_case_numb_col			= 11
const case_numb_correct_col		= 12
const testing_info_col			= 13
const running_error_col			= 14


total_excel_row = 2
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

ObjExcel.Cells(1, script_run_date_time_col).Value  = "Script Run Date & time"
ObjExcel.Cells(1, script_user_col).Value  		= "Script User"
ObjExcel.Cells(1, file_name_col).Value  		= "File Name"
ObjExcel.Cells(1, confrim_numb_col).Value  		= "Confirmation Number"
ObjExcel.Cells(1, appl_date_and_time).Value  	= "Application Date & Time"
ObjExcel.Cells(1, script_run_time).Value  		= "Script Run Seconds"
ObjExcel.Cells(1, panel_title_end).Value  		= "PANEL TITLE at Script End"
ObjExcel.Cells(1, search_type_col).Value  		= "Search Type"
ObjExcel.Cells(1, search_found_col).Value  		= "Search Success"
ObjExcel.Cells(1, search_notes_col).Value  		= "Search Notes"
ObjExcel.Cells(1, xml_case_numb_col).Value  	= "XML Case Number"
ObjExcel.Cells(1, case_numb_correct_col).Value 	= "Case Number Correct"
ObjExcel.Cells(1, testing_info_col).Value  		= "Test Notes"
ObjExcel.Cells(1, running_error_col).Value  	= "Running Error"

for i = 1 to 14
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
next


'defining the assignment folder and setting the file paths
testing_log_folder = t_drive & "\Eligibility Support\Assignments\Script Testing Logs"
Set objFolder = objFSO.GetFolder(testing_log_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

'Looking at each txt file in the assignments folder to capture Expedited Determination information
For Each objFile in colFiles																'looping through each file
    this_file_path = objFile.Path												'identifying the current file
	this_file_name = objFile.Name
	this_file_created_date = objFile.DateCreated								'Reading the date created

	If DateDiff("d", this_file_created_date, date) <> 0 Then						'we are only pulling information for cases that were reviewed yesterday at this time.
	    'Setting the object to open the text file for reading the data already in the file
	    Set objTextStream = objFSO.OpenTextFile(this_file_path, ForReading)

	    'Reading the entire text file into a string
	    every_line_in_text_file = objTextStream.ReadAll

	    exp_det_details = split(every_line_in_text_file, vbNewLine)					'creating an array of all of the information in the TXT files

		For Each text_line in exp_det_details										'read each line in the file
			If Instr(text_line, ":") <> 0 Then
				line_info = split(text_line, ":")								'creating a small array for each line. 0 has the header and 1 has the information
				line_info(0) = trim(line_info(0))
				'here we add the information from TXT to Excel


				If line_info(0) = "SCRIPT Run Date and Time" Then
					script_run_info = ""
					for i = 1 to UBound(line_info)
						script_run_info = script_run_info & ":" & line_info(i)
					next
					if left(script_run_info, 1) = ":" Then script_run_info = right(script_run_info, len(script_run_info)-1)
					ObjExcel.Cells(total_excel_row, script_run_date_time_col).Value  = script_run_info
				End If
				If line_info(0) = "Script run be" Then ObjExcel.Cells(total_excel_row, script_user_col).Value  = line_info(1)
				If line_info(0) = "File Name Selected" Then ObjExcel.Cells(total_excel_row, file_name_col).Value  = line_info(1)
				If line_info(0) = "Confirmation Number" Then ObjExcel.Cells(total_excel_row, confrim_numb_col).Value  = line_info(1)
				If line_info(0) = "APPL Date" Then
					appl_date_info = ""
					for i = 1 to UBound(line_info)
						appl_date_info = appl_date_info & ":" & line_info(i)
					next
					if left(appl_date_info, 1) = ":" Then appl_date_info = right(appl_date_info, len(appl_date_info)-1)
					ObjExcel.Cells(total_excel_row, appl_date_and_time).Value  = appl_date_info
				End If
				If line_info(0) = "Length of script run" Then ObjExcel.Cells(total_excel_row, script_run_time).Value  = line_info(1)
				If line_info(0) = "Panel Title at the End" Then ObjExcel.Cells(total_excel_row, panel_title_end).Value  = line_info(1)
				If line_info(0) = "Search Type" Then ObjExcel.Cells(total_excel_row, search_type_col).Value  = line_info(1)
				If line_info(0) = "Was the search found" Then ObjExcel.Cells(total_excel_row, search_found_col).Value  = line_info(1)
				If line_info(0) = "Search Notes" Then ObjExcel.Cells(total_excel_row, search_notes_col).Value  = line_info(1)
				' No Case Number was found on the XML.
				If line_info(0) = "Case Number from Form" Then ObjExcel.Cells(total_excel_row, xml_case_numb_col).Value  = line_info(1)
				If line_info(0) = "Does this Case Number appear to be accurate" Then ObjExcel.Cells(total_excel_row, case_numb_correct_col).Value  = line_info(1)
				If line_info(0) = "Testing Information" Then ObjExcel.Cells(total_excel_row, testing_info_col).Value  = line_info(1)
				If line_info(0) = "Running Error" Then ObjExcel.Cells(total_excel_row, running_error_col).Value  = line_info(1)

			Else
				If text_line = "No Case Number was found on the XML." Then ObjExcel.Cells(total_excel_row, xml_case_numb_col).Value = "None"
			End If
		Next

		total_excel_row = total_excel_row + 1										'advance to the next row

		objTextStream.Close						'we are done with this file, so we must close the access

		Dim oTxtFile
		' With (CreateObject("Scripting.FileSystemObject"))
		' 	'If the file exists in the archive, we we will delete the version in the archive so the one from the main file can be placed in archive
		' 	If .FileExists(txt_file_archive_path & "\" & this_file_name & ".txt") Then
		' 		objFSO.DeleteFile(txt_file_archive_path & "\" & this_file_name & ".txt")		'deleting the TXT file because hgave the information
		' 	End If
		' End With
		' On error resume next
		' If Err.Number <> 0 Then MsgBox "FILE IS DUPLICATE ???" & vbCr & "this_file_path - " & this_file_path & vbCr & "archive pather - " & txt_file_archive_path & "\" & this_file_name & ".txt"
		' On Error Goto 0
	End If
Next

Call script_end_procedure("Do we have a report?")