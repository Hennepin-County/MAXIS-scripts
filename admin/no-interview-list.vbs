'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - No Interview List.vbs"
start_time = timer
STATS_counter = 0			     'sets the stats counter at one
STATS_manualtime = 	90			 'manual run time in seconds
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
call changelog_update("07/06/2026", "Initial version.", "Dave Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

If db_full_string = "" Then
db_provider = "SQLOLEDB.1"
db_data_source = "hssqlpw202"
db_catalog = "BlueZone_Statistics"
db_security = "SSPI"
db_translate = "False"

'string to use for database calls in scripts.
db_full_string = "Provider = " & db_provider & ";Data Source= " & db_data_source & ";Initial Catalog= " & db_catalog & "; Integrated Security=" & db_security & ";Auto Translate=" & db_translate & ";"
End If

Do
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 181, 55, "No Interview List"
		ButtonGroup ButtonPressed
			OkButton 120, 10, 50, 15
			CancelButton 120, 30, 50, 15
		Text 10, 10, 100, 30, "This script generates a list of cases with no interview completed and SNAP pending."
    EndDialog

	dialog Dialog1
	cancel_without_confirmation

	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
Call check_for_MAXIS(False)


	objSQL = "SELECT * From [BlueZone_Statistics].[ES].[ES_OnDemandCashAndSnap] as OD Inner Join [BlueZone_Statistics].[ES].[ES_OnDemanCashAndSnapBZProcessed] as BZ ON Cast(OD.CaseNumber as int) = cast(BZ.CaseNumber as int) Where BZ.SnapStatus = 'Pending'"


	Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open db_full_string
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "Caseload #"
ObjExcel.Cells(1, 2).Value = "Case Number"
ObjExcel.Cells(1, 3).Value = "Appl Date"
ObjExcel.Cells(1, 4).Value = "Interview Date"
ObjExcel.Cells(1, 5).Value = "Days Pending"
FOR i = 1 to 5	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
NEXT
ExcelRow = 2	'starting row for data entry


Do While NOT objRecordSet.Eof
	Objexcel.Cells(ExcelRow, 1).Value = objRecordSet.Fields("WorkerID").Value
	Objexcel.Cells(ExcelRow, 2).Value = objRecordSet.Fields("CaseNumber").Value
	Objexcel.Cells(ExcelRow, 3).Value = objRecordSet.Fields("ApplDate").Value
	Objexcel.Cells(ExcelRow, 4).Value = objRecordSet.Fields("InterviewDate").Value
	Objexcel.Cells(ExcelRow, 5).Value = objRecordSet.Fields("DaysPending").Value
	ExcelRow = ExcelRow + 1
	objRecordSet.MoveNext		'go to the next case
Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close

Script_end_procedure("Success! Your list has been generated.")
