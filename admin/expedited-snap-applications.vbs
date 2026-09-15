' ================================================================
' EXPEDITED SNAP APPS REPORT
' ================================================================

'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - EXPEDITED SNAP APPLICATIONS.vbs"
start_time = timer
STATS_counter = 1                  	'sets the stats counter at one
STATS_manualtime = 0              'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = True
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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================
'Global variables for local runs


If db_full_string = "" Then
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-Scripts\locally-installed-files\SETTINGS - GLOBAL VARIABLES.vbs")
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
End If

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")

call changelog_update("09/15/2026", "Initial version.", "Dave Courtright, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Limit the script only to certain users
Allow_use = False
If user_ID_for_validation = "CALO001" Then allow_use = True
If user_ID_for_validation = "DACO003" Then allow_use = True
If user_ID_for_validation = "ASRE002" Then allow_use = True
If user_ID_for_validation = "TRFA001" Then allow_use = True
If user_ID_for_validation = "SBegleyMay" Then allow_use = True

If user_ID_for_validation = "WFV833" Then allow_use = True 'Ryan
If user_ID_for_validation = "WFU161" Then allow_use = True 'Brooke

If allow_use = False Then
	critical_error_msgbox = MsgBox ("You are not authorized to run this script. Please contact a scripts administrator with any questions.", vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
	StopScript
End If

appl_entry_date = DateAdd("d", -1, Date)

Do While Weekday(appl_entry_date, vbMonday) > 5
	appl_entry_date = DateAdd("d", -1, appl_entry_date)
Loop

appl_entry_date = "" & appl_entry_date
Dialog1 = ""

BeginDialog Dialog1, 0, 0, 210, 80, "Expedited SNAP Applications"
  Text 10, 15, 85, 10, "APPL entry Date:"
  EditBox 100, 12, 95, 15, appl_entry_date
  OkButton 95, 50, 50, 15
  CancelButton 150, 50, 50, 15
EndDialog

Do
    err_msg = ""
    Dialog Dialog1
    cancel_confirmation
    If Not IsDate(appl_entry_date) Then err_msg = err_msg & vbcr & "Please enter a valid date."
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
Loop Until err_msg = ""


appl_entry_date = CDate(appl_entry_date)

Const adCmdText = 1
Const adParamInput = 1
Const adDBDate = 133



sql_query = "Select WorkerID, CaseNumber, ApplDate, UpdateDate from Es.ES_CasesPending Where IsExpSNAP = 1 And UpdateDate = '" & appl_entry_date & "'"

Set objConn = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConn.Open db_full_string
	objRecordSet.Open sql_query, objConn

Set excel_app = CreateObject("Excel.Application")
excel_app.Visible = True

Set workbook = excel_app.Workbooks.Add()
Set worksheet = workbook.Worksheets(1)
worksheet.Name = "Expedited SNAP"

For col = 0 To objRecordSet.Fields.Count - 1
	worksheet.Cells(1, col + 1).Value = objRecordSet.Fields(col).Name
Next

row = 2
Do Until objRecordSet.EOF
	For col = 0 To objRecordSet.Fields.Count - 1
		worksheet.Cells(row, col + 1).Value = objRecordSet.Fields(col).Value
	Next
	row = row + 1
	objRecordSet.MoveNext
Loop

worksheet.UsedRange.Columns.AutoFit
objRecordSet.Close
objConn.Close

Set objRecordSet = Nothing
Set objConn = Nothing

end_msg = "The report has been generated. Please check Excel for the results."
script_end_procedure_with_error_report(end_msg)