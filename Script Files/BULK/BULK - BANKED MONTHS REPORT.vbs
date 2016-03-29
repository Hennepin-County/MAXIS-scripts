'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - MONTHLY BANKED MONTHS DATA GATHER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1                          'sets the stats counter at one
'STATS_manualtime = ***                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'FUNCTIONS that are currently not in the FuncLib that are used in this script----------------------------------------------------------------------------------------------------
'Veronicas function that allows the user to search for a local file instead of having the file location hard coded into the script'
Function File_Selection_System_Dialog(file_selected)
    'Creates a Windows Script Host object
    Set wShell=CreateObject("WScript.Shell")

    'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

    'Creates the file_selected variable from the exit
    file_selected = oExec.StdOut.ReadLine
End function

'We will be writing a person note as we gather Banked Months Data so we need these functions'
Function write_editbox_in_person_note(x, y) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
  variable_array = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in variable_array
    EMGetCursor row, col
    If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col
    If (row < 18 and col + (len(x)) >= 80) then EMSendKey "<newline>" & space(5)
    If (row = 5 and col = 3) then EMSendKey space(5)
    EMSendKey x & " "
    If right(x, 1) = ";" then
      EMSendKey "<backspace>" & "<backspace>"
      EMGetCursor row, col
      If row = 18 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSendKey space(5)
      Else
        EMSendKey "<newline>" & space(5)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_new_line_in_person_note(x)
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80 + 1 ) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

'FUNCTION create_dialog(month_list, report_month_dropdown)
	'DIALOGS-----------------------------------------------------------------------------------------------------------
	'NOTE: droplistbox for scenario list must be: ["select one..." & scenario_list] in order to be dynamic
'	BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 211, 70, "SNAP Banked Month Reporting Dialog"
'	  DropListBox 65, 25, 140, 15, "select one..." & month_list, report_month_dropdown
'	  ButtonGroup ButtonPressed
'		OkButton 100, 45, 50, 15
'		CancelButton 155, 45, 50, 15
'	  Text 5, 10, 190, 10, "Select the month that you are creating the report for."
'	  Text 5, 30, 55, 10, "Month to Report:"
'	EndDialog
'
'	DIALOG SNAP_Banked_Month_Report_Dialog
'END FUNCTION

'DIALOGS----------------------------------------------------------------------------------------------------
'BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 211, 70, "SNAP Banked Month Reporting Dialog"
'  DropListBox 65, 25, 140, 15, "select one..." & month_list, report_month_dropdown
'  ButtonGroup ButtonPressed
'    OkButton 100, 45, 50, 15
'    CancelButton 155, 45, 50, 15
'  Text 5, 10, 190, 10, "Select the month that you are creating the report for."
'  Text 5, 30, 55, 10, "Month to Report:"
'EndDialog

EMConnect ""		'connecting to MAXIS

MsgBox "You need to open the Excel File that has the list of clients reported as using a banked month for the month being reported." & _
  VBNewLine & VBNewLine & "Be sure your spreadsheet is in the correct format." 'Notice to the user that a finder window will open for them to search for their list of client that have used banked months'
Call File_Selection_System_Dialog(list_reported_banked_month_clients)  'References the function above to have the user seach for their file'
call excel_open(list_reported_banked_month_clients, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

'This reads every worksheet name in the selected excel file and creates a list for the drop down in the dialog'
For Each objWorkSheet In objWorkbook.Worksheets
	month_list = month_list & chr(9) & objWorkSheet.Name
Next

'This is the dialog for the user to select which month (or worksheet) of data they are going to generate a report for'
BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 211, 70, "SNAP Banked Month Reporting Dialog"
  DropListBox 65, 25, 140, 15, "select one..." & month_list, report_month_dropdown
  ButtonGroup ButtonPressed
	OkButton 100, 45, 50, 15
	CancelButton 155, 45, 50, 15
  Text 5, 10, 190, 10, "Select the month that you are creating the report for."
  Text 5, 30, 55, 10, "Month to Report:"
EndDialog

'Runs the dialog'
Do
	Dialog SNAP_Banked_Month_Report_Dialog
	cancel_confirmation
Loop until report_month_dropdown <> "select one..."

'This assigns a footer month and year based on the worksheet names selected in the dropdown from the dialog'
Select Case report_month_dropdown
Case "January 2016"
	footer_month = "01"
	footer_year = "16"
Case "February 2016"
	footer_month = "02"
	footer_year = "16"
End Select
MsgBox footer_month & "/" & footer_year

'Sets up the array to store all the information for each client'
Dim Banked_Month_Client_Array ()
ReDim Banked_Month_Client_Array (10, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num = 1
Const clt_pmi = 2
Const memb_num = 3
Const clt_name = 4
Const clt_first_name = 5
Const clt_last_name = 6
Const comments = 7
Const abawd_used = 8
Const second_abawd_used = 9
Const send_to_DHS = 10


'Now the script adds all the clients on the excel list into an array
excel_row = 5 're-establishing the row to start checking the members for
entry_record = 0
Do                                                                                          'Loops until there are no more cases in the Excel list
	case_number = objExcel.cells(excel_row, 4).Value          're-establishing the case numbers
	If case_number = "" then exit do
	case_number = trim(case_number)
	client_first_name = objExcel.cells(excel_row, 3).Value
	client_last_name = objExcel.cells(excel_row, 2).Value             're-establishing the client name
	client_first_name = UCase(trim(client_first_name))
	client_last_name = UCase(trim(client_last_name))
'Adding client information to the array'
	ReDim Preserve Banked_Month_Client_Array(send_to_DHS, entry_record)
	Banked_Month_Client_Array (case_num, entry_record) = case_number
	Banked_Month_Client_Array (clt_last_name, entry_record) = client_last_name
	Banked_Month_Client_Array (clt_first_name, entry_record) = client_first_name
	Banked_Month_Client_Array (clt_name, entry_record) = client_first_name & " " & client_last_name
	Banked_Month_Client_Array (comments, entry_record) = objExcel.cells(excel_row, 6).Value

	'MsgBox client_first_name & " " & client_last_name & VBNewLine & Banked_Month_Client_Array (clt_name,entry_record)
	entry_record = entry_record + 1
	excel_row = excel_row + 1
Loop

'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document
'and don't want the script to be confused
objExcel.Quit

'Now we will get PMI and Member Number for each client on the array.'
For item = 0 to UBound(Banked_Month_Client_Array, 2)
	case_number = Banked_Month_Client_Array(case_num,item)	'Case number is set for each loop as it is used in the FuncLib functions'
	Call navigate_to_MAXIS_screen("STAT", "MEMB")						'Finding client information on STAT MEMB'
	DO
		EMReadScreen membs_on_case, 2, 2, 78									'Reads how many MEMB panels are on each case'
		membs_on_case = trim(membs_on_case)
		membs_on_case = right ("00" & membs_on_case,2)				'Adds a leading 0 to the number of members and makes it a 2 didgit value'
		EMReadScreen last_name, len(Banked_Month_Client_Array(clt_last_name, item)), 6, 30	'Reads the last name of the client from the Member panel - it only reads the number of characters that were in the initial excel list - this prevents the "JR" suffixes from causing a match failure
		EMReadScreen first_name, len(Banked_Month_Client_Array(clt_first_name,item)), 6, 63	'Reads the first name of the client from the Member panel - it only reads the number of characters that were in the initial excel list - this prevents the "JR" suffixes from causing a match failure
		'last_name = Replace(last_name,"_","")
		'first_name = Replace(first_name,"_","")
		name_of_client = trim(first_name & " " & last_name)                       'creates a full name
			If Banked_Month_Client_Array(clt_name,item) = name_of_client then				'Compares the name in MAXIS with the name on the excel list'
				EMReadScreen memb_number, 2, 4, 33																		'If they match it will read the member number and PMI and add the information to the array'
				EMReadScreen PMI_number, 9,  4, 46
				Banked_Month_Client_Array (clt_pmi, item) = replace(trim(PMI_number), "_", "")
				Banked_Month_Client_Array (memb_num, item) = memb_number
				exit do
			ELSEIF membs_on_case = "01" Then										'This identifies if there is only one person on the case AND the list name and MAXIS name do not match'
				ConfirmAdd = MsgBox ("This is the only client listed on this case" & VBNewLine & "Click 'Yes' to add this client to the Banked Month report as having used an ABAWD Banked Month", vbYesNo + vbQuestion)
				If ConfirmAdd = vbYes Then												'If there is only one person on the case AND the names don't match, this message box asks the user if this is the person.
					EMReadScreen memb_number, 2, 4, 33							'If they say YES then the script will add the PMI and Member number to the array'
					EMReadScreen PMI_number, 9,  4, 46
					Banked_Month_Client_Array (clt_pmi, item) = replace(trim(PMI_number), "_", "")
					Banked_Month_Client_Array (memb_num, item) = memb_number
					exit do
				ElseIF vbNo Then																	'If they say NO then it will mark the client as NOT being reported as using a banked month'
					Banked_Month_Client_Array (send_to_DHS, item) = FALSE
					Exit Do
				End If
			ELSE																								'IF the name doesn't match and there is more than one member in the case, the script will transmit through the members and compare all the names
				EMReadScreen memb_number, 2, 4, 33
				transmit
				EMReadScreen next_memb_number, 2, 4, 33
				IF memb_number = next_memb_number Then						'identifies if the script has reached the end of the memnbers'
					Banked_Month_Client_Array (send_to_DHS, item) = FALSE		'If it has reached the end of the members and there is still no match the script will not add this person to the DHS document to report'
					MsgBox "There is no client listed on this case that matches the client reported as using a banked month. You will need to add this person to the DHS Reporting Log Manually" & _
					  vbNewLine & vbNewLine & Banked_Month_Client_Array(clt_name, item), vbExclamation		'This will alert the worker that a match could not be found in this case and they will need to manually find this person'
					Exit Do
				End If
			END IF
			'Msgbox Banked_Month_Client_Array(clt_name,item) & vbnewline & name_of_client
	LOOP until Banked_Month_Client_Array(clt_name, item) = name_of_client OR Banked_Month_Client_Array(send_to_DHS, item) = FALSE  'Ends the loop if a match is found or script has determined to NOT report on this person'

Next				'Goes to the next Array item to compare'

For i = 0 to Ubound(Banked_Month_Client_Array,2)
	MsgBox "Case # " & Banked_Month_Client_Array (case_num, i) & vbNewLine & "PMI: " & Banked_Month_Client_Array(clt_pmi,i) & vbNewLine & "Memb " & Banked_Month_Client_Array(memb_num, i) & vbNewLine & "Name: " & Banked_Month_Client_Array(clt_name, i) & vbNewLine & "Name again: " & Banked_Month_Client_Array (clt_first_name, i) & " " & Banked_Month_Client_Array(clt_last_name, i) & vbNewLine & "Comments: " & Banked_Month_Client_Array(comments, i)
Next

'Writing to the DHS tracking sheet
'MsgBox "Selct the file of the Excel Spreadsheet you submit to DHS" & _
'	VBNewLine & VBNewLine & "Be sure your spreadsheet is in the correct format."
'Call File_Selection_System_Dialog(list_reported_banked_month_clients)
'call excel_open(list_reported_banked_month_clients, True, True, ObjExcel, objWorkbook)

'For Each objWorkSheet In objWorkbook.Worksheets
'	month_list = month_list & chr(9) & objWorkSheet.Name
'Next


'BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 211, 70, "SNAP Banked Month Reporting Dialog"
'  DropListBox 65, 25, 140, 15, "select one..." & month_list, report_month_dropdown
'  ButtonGroup ButtonPressed
'	OkButton 100, 45, 50, 15
'	CancelButton 155, 45, 50, 15
'  Text 5, 10, 190, 10, "Select the month that you are creating the report for."
'  Text 5, 30, 55, 10, "Month to Report:"
'EndDialog

'Do
'	Dialog SNAP_Banked_Month_Report_Dialog
'	cancel_confirmation
'Loop until report_month_dropdown <> "select one..."

script_end_procedure("")
