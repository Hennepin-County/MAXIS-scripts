'Required for statistical purposes===============================================================================
name_of_script = "BULK - ADD PNOTE TO BANKED MONTHS CASES.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'FUNCTIONS that are currently not in the FuncLib that are used in this script----------------------------------------------------------------------------------------------------
Function File_Selection_System_Dialog(file_selected)
    'Creates a Windows Script Host object
    Set wShell=CreateObject("WScript.Shell")

    'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

    'Creates the file_selected variable from the exit
    file_selected = oExec.StdOut.ReadLine
End function

'-------THIS FUNCTION ALLOWS THE USER TO PICK AN EXCEL FILE---------
Function BrowseForFile()
    Dim shell : Set shell = CreateObject("Shell.Application")
    Dim file : Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "Computer")
	IF file is Nothing THEN 
		script_end_procedure("The script will end.")
	ELSE
		BrowseForFile = file.self.Path
	END IF
End Function

'Dialog----------------------------------------------------------------------------------------------------
BeginDialog update_banked_month_status_dialog, 0, 0, 191, 60, "Dialog"
  DropListBox 80, 10, 105, 15, "Select one..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", month_selection
  ButtonGroup ButtonPressed
    OkButton 80, 35, 50, 15
    CancelButton 135, 35, 50, 15
  Text 5, 15, 70, 10, "Update status month:"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog update_banked_month_status_dialog
		If ButtonPressed = 0 then StopScript
		If month_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the status month to update."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					
	
'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

'starts adding phone numbers at row selected
Select Case month_selection
Case "January"
	MAXIS_footer_month = "01"
	MAXIS_footer_year = "16"
	report_date = "01/16"
	'excel_worksheet = "January 2016"
Case "February"
	MAXIS_footer_month = "02"
	MAXIS_footer_year = "16"
	report_date = "02/16"
	'excel_worksheet = "February 2016"
Case "March"
	MAXIS_footer_month = "03"
	MAXIS_footer_year = "16"
	report_date = "03/16"
	'excel_worksheet = "March 2016"
Case "April"
	MAXIS_footer_month = "04"
	MAXIS_footer_year = "16"
	report_date = "04/16"
	'excel_worksheet = "April 2016"
Case "May"
	MAXIS_footer_month = "05"
	MAXIS_footer_year = "16"
	report_date = "05/16"
	'excel_worksheet = "May 2016"
Case "June"
	MAXIS_footer_month = "06"
	MAXIS_footer_year = "16"
	report_date = "06/16"
	'excel_worksheet = "June 2016"
Case "July"
	MAXIS_footer_month = "07"
	MAXIS_footer_year = "16"
	report_date = "07/16"
	'excel_worksheet = "July 2016"
Case "August"
	MAXIS_footer_month = "08"
	MAXIS_footer_year = "16"
	report_date = "08/16"
	'excel_worksheet = "August 2016"
Case "September"
	MAXIS_footer_month = "09"
	MAXIS_footer_year = "16"
	report_date = "09/16"
	'excel_worksheet = "September 2016"
Case "October"
	MAXIS_footer_month = "10"
	MAXIS_footer_year = "16"
	report_date = "10/16"
	'excel_worksheet = "October 2016"
Case "November"
	MAXIS_footer_month = "11"
	MAXIS_footer_year = "16"
	report_date = "11/16"
	'excel_worksheet = "November 2016"
Case "December"
	MAXIS_footer_month = "12"
	MAXIS_footer_year = "16"
	report_date = "12/16"
	'excel_worksheet = "December 2016"
End Select

'function to call up a local/network file
DO 
	DO
		'file_location = InputBox("Please enter the file location.")
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open(BrowseForFile)
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
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in		

'This reads every worksheet name in the selected excel file and creates an array that will be used to determine which month is being reported'
For Each objWorkSheet In objWorkbook.Worksheets
	DHS_report_month_list = DHS_report_month_list & "~" & objWorkSheet.Name
	If left(DHS_report_month_list, 1) = "~" then DHS_report_month_list = right(DHS_report_month_list, len(DHS_report_month_list) - 1)
	DHS_report_month_array = Split(DHS_report_month_list,"~")
Next

'The user already selected the month in the initial excel sheet - this was used to set the footer month.'
'Now the footer month is used to select the right worksheet in the DHS report to match'
Select Case MAXIS_footer_month
Case "01"
	DHS_report_month = DHS_report_month_array(0)	'January, arrays start at 0'
Case "02"
	DHS_report_month = DHS_report_month_array(1)	'February'
Case "03"
	DHS_report_month = DHS_report_month_array(2)	'March'
Case "04"
	DHS_report_month = DHS_report_month_array(3)	'April'
Case "05"
	DHS_report_month = DHS_report_month_array(4)	'May'
Case "06"
	DHS_report_month = DHS_report_month_array(5)	'June'
Case "07"
	DHS_report_month = DHS_report_month_array(6)	'July'
Case "08"
	DHS_report_month = DHS_report_month_array(7)	'August'
Case "09"
	DHS_report_month = DHS_report_month_array(8)	'September'
Case "10"
	DHS_report_month = DHS_report_month_array(9)	'October'
Case "11"
	DHS_report_month = DHS_report_month_array(10)	'November'
Case "12"
	DHS_report_month = DHS_report_month_array(11)	'December'
End Select

'Activates the selected worksheet
objExcel.worksheets(DHS_report_month).Activate

'Sets up the array to store all the information for each client'
Dim PNOTE_array ()
ReDim PNOTE_array (2, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num = 1			'Each of the case numbers will be stored at this position'
Const clt_PMI  = 2			'PMI stored

'creating an array of cases to update
excel_row = 2
entry_record = 0
Do 
	MAXIS_case_number = objExcel.cells(excel_row, 2).value
	MAXIS_case_number = trim(MAXIS_case_number)
	PMI_number = objExcel.cells(excel_row, 3).value
	PMI_number = trim(PMI_number)
	IF MAXIS_case_number = "" then exit do
	'Adding client information to the array'
	ReDim Preserve PNOTE_array(2, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	PNOTE_array	(case_num, entry_record) = MAXIS_case_number		'The client information is added to the array'
	PNOTE_array (clt_PMI,  entry_record) = PMI_number
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	excel_row = excel_row + 1
LOOP

'Function to ensure that we're in the correct footer month
MAXIS_footer_month_confirmation	

For item = 0 to UBound(PNOTE_array, 2)
	MAXIS_case_number = PNOTE_array(case_num, item)				'Case number is set for each loop as it is used in the FuncLib functions'
	PMI_number = PNOTE_array(clt_PMI, item)
	
	IF MAXIS_case_number <> "" then 
		EMWriteScreen "________", 18, 43					'clears case number
		EMWriteScreen MAXIS_case_number, 18, 43
		
		'checking the PMI number on the MEMB panel
		Call navigate_to_MAXIS_screen("STAT", "MEMB")
		Do 
			EMReadScreen MAXIS_PMI_number, 7, 4, 46
			MAXIS_PMI_number = trim(MAXIS_PMI_number)
			IF MAXIS_PMI_number = PMI_number then 
				EMReadScreen member_number, 2, 4, 33
				exit do
			else 
				transmit				'to go to the next member on MEMB
				EMReadScreen last_member, 5, 24, 2	
			END IF 
		Loop until last_member = "ENTER" 
		IF MAXIS_PMI_number <> PMI_number then msgbox MAXIS_case_number & " did not match the the PMI number. Please update PNOTES manually. PMI number: " & PMI_number
		
		CALL navigate_to_MAXIS_screen("STAT", "WREG")
		EMWriteScreen member_number, 20, 76
		transmit
		'going into edit mode for person notes and adding person notes for the selected month if they do not exist already.
		PF5
		EMReadScreen PNOTE_check, 4, 2, 46
		If PNOTE_check <> "SCRN" then 
			msgbox MAXIS_case_number & " cannot be updated. Review case, and add person note manually if applicable."
		ELSE
			EMreadscreen edit_mode_required_check, 6, 5, 3		'if not person not exists, person note goes directly into edit mode
			If edit_mode_required_check = "      " then 
				EMWriteScreen "Banked Month Used " & report_date, 5, 3
				EMWriteScreen "Case has been counted and reported to DHS.", 6, 3 
			ElseIF edit_mode_required_check <> "      " then 	
				'creating a Do loop to ensure that duplicate person notes are not being made
				PNOTE_row = 5		'establishes the row to start searching the Person notes from
				Do
					EMReadScreen counted_banked_month, 12, PNOTE_row, 31
					If counted_banked_month = "Banked Month" then EMReadScreen abawd_counted_months_string, 5, PNOTE_row, 49
					If abawd_counted_months_string = report_date then exit do	'if person note has already been made for the report date, then does not person note
					PNOTE_row = PNOTE_row + 1	'adds incremental to row to search
				LOOP until PNOTE_row = 18
				If PNOTE_row = 18 then 
					PF9
					EMWriteScreen "Banked Month Used " & report_date, 5, 3
					EMWriteScreen "Case has been counted and reported to DHS.", 6, 3 
				END IF
			END IF
		END IF 	
		STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
		Do
			EMSendKey "<PF3>"
    		EMWaitReady 0, 0
    		EMReadScreen SELF_check, 4, 2, 50
		LOOP until SELF_check = "SELF"	
	End if
Next

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! The person notes have been updated!")