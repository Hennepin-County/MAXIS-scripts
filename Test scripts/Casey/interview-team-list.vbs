'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
'These are the constants that we need to create tables in Excel
Const xlSrcRange = 1
Const xlYes = 1

If NOT IsArray(interviewer_array) Then
	Dim tester_array()
	ReDim tester_array(0)
	Dim interviewer_array()
	ReDim interviewer_array(0)
	tester_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")

	Set fso_command = run_another_script_fso.OpenTextFile(tester_list_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
End If

'Creating the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first row with header information
ObjExcel.Cells(1, 1).Value = "FULL NAME"
ObjExcel.Cells(1, 2).Value = "FIRST NAME"
ObjExcel.Cells(1, 3).Value = "LAST NAME"
ObjExcel.Cells(1, 4).Value = "EMAIL"
ObjExcel.Cells(1, 5).Value = "WINDOWS LOGON"
ObjExcel.Cells(1, 6).Value = "X-NUMBER"
ObjExcel.Cells(1, 7).Value = "TRAINER"

excel_row = 2
For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
	ObjExcel.Cells(excel_row, 1).Value = worker.interviewer_full_name
	ObjExcel.Cells(excel_row, 2).Value = worker.interviewer_first_name
	ObjExcel.Cells(excel_row, 3).Value = worker.interviewer_last_name
	ObjExcel.Cells(excel_row, 4).Value = worker.interviewer_email
	ObjExcel.Cells(excel_row, 5).Value = worker.interviewer_id_number
	ObjExcel.Cells(excel_row, 6).Value = worker.interviewer_x_number
	ObjExcel.Cells(excel_row, 7).Value = worker.interview_trainer

	excel_row = excel_row+1
Next

For col_to_autofit = 1 to 7
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

table_range = "A1:G" & excel_row
ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table_range, xlYes).Name = "InterviewerTable"
ObjExcel.ActiveSheet.ListObjects("InterviewerTable").TableStyle = "TableStyleMedium4"


StopScript