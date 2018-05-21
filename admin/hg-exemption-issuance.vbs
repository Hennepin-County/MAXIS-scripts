'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - HG expansion issuance.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "420"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/07/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			BeginDialog HG_issuance_dialog, 0, 0, 266, 125, "HG expansion issuance dialog"
			    ButtonGroup ButtonPressed
    			PushButton 200, 60, 50, 15, "Browse...", select_a_file_button
    			OkButton 145, 105, 50, 15
    			CancelButton 200, 105, 50, 15
  				EditBox 15, 60, 180, 15, HG_path
  				GroupBox 10, 5, 250, 95, "Using the script"
  				Text 15, 20, 235, 35, "This script should be used when DHS provides your county with a list of recipeints that are eligible for the HG expansion. It will gather the case numbers from the list, and gather the issaunce information for 08/16-02/17. This is the window of time for manual issuances."
  				Text 15, 80, 230, 15, "Select the Excel file that contains the HG inforamtion by selecting the 'Browse' button, and finding the file."
			EndDialog
			
			err_msg = ""
			Dialog HG_issuance_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If HG_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(HG_path, ".xlsx") 'allows the user to select the file'
			End If
			If HG_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(HG_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Now the script adds all the cases on the applicable spreadsheet into an array
excel_row = 2 	're-establishing the row to start checking - excel row 9 is the row that DHS has assinged as the 1st column with a case number
entry_record = 0

Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 1).Value          'estbalishing 
	If MAXIS_case_number = "" then exit do
	MAXIS_case_number = trim(MAXIS_case_number)
	
	all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
	If MAXIS_case_number <> "" then case_number_list = case_number_list & MAXIS_case_number & "," 

	entry_record = entry_record + 1			'This increments to the next entry in the array
	excel_row = excel_row + 1
	'blanking out variables
	MAXIS_case_number = ""
Loop

'msgbox entry_record

If entry_record = 0 then script_end_procedure("No cases have been found on this list for your county. The script wil now end.")

'Closes the list of HG recipients since we don't need this anymore
objExcel.Quit

'ARRAY business----------------------------------------------------------------------------------------------------
'Sets up the array to store all the information for each client'
Dim HG_array ()
ReDim HG_array (6, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_number	= 1			'Each of the case numbers will be stored at this position'
'Const aug_2016		= 2
'Const sept_2016	= 3	
'Const oct_2016		= 4
'Const nov_2016    	= 5
'Const dec_2016    	= 6
Const jan_2017    	= 2
Const feb_2017		= 3
Const march_2017	= 4
Const april_2017	= 5
Const may_2017		= 6

case_count = -1			'establishes the value of the case_count for the array. Since array values start at 0, count starts at -1

'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
case_number_list = trim(case_number_list)
If right(case_number_list, 1) = "," then case_number_list = left(case_number_list, len(case_number_list) - 1)
case_numbers_array = split(case_number_list, ",")

'msgbox case_number_list
For each MAXIS_case_number in case_numbers_array
	case_count = case_count + 1
	IF MAXIS_case_number = "" then exit for 
	'msgbox MAXIS_case_number
	Call navigate_to_MAXIS_screen("MONY", "INQX")
	
	EMWriteScreen "12", 6, 38		'entering 08/16 as this is the 1st month that we need to Check
	EMWriteScreen "16", 6, 41
	EMWriteScreen "05", 6, 53		'entering 02/17 as this is the last month we need to check
	EMWriteScreen "17", 6, 56
	EMWriteScreen "x", 10, 5		'selecting MFIP
	transmit
	
	'creating an array of issuance months to fill in the Excel list from INQD
	'issuance_months_array = array("08/16", "09/16", "10/16", "11/16", "12/16", "01/17", "02/17")
	issuance_months_array = array("01/17", "02/17", "03/17", "04/17", "05/17")
	'searching for the housing grant issued on the INQX/INQD screen(s)
	For each issuance_month in issuance_months_array 		'For next searches issuances for all rolling 12 months
		DO
			row = 6				'establishing the row to start searching for issuance'
			DO
				EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
				If trim(housing_grant) = "" then exit do		'exits the do loop once the end of the issuances is reached
				IF housing_grant = "HG" then
					'reading the housing grant information
					EMReadScreen HG_amt_issued, 7, row, 40
					EMReadScreen HG_month, 2, row, 73
					EMReadScreen HG_year, 2, row, 79
					INQD_issuance_month = HG_month & "/" & HG_year		'creates a new varible for HG month and year
					If issuance_month = INQD_issuance_month then 		'if the issuance found matches the issuance month then
						HG_amt_issued = trim(HG_amt_issued)				'trims the HG amt issued variable
						'msgbox issuance_month & vbcr & HG_amt_issued
						'msgbox "exit 1st do"
						exit do
					ELSE
						HG_amt_issued = "0"
					END IF
				END IF
				row = row + 1											'adds one row to search the next row
			Loop until row = 18	
													'repeats until the end of the page
			If row = 18 then
				EMReadScreen last_page_check, 21, 24, 2
				'msgbox last_page_check
			 	IF trim(last_page_check) = "THIS IS THE 1ST PAGE" then 
					exit do
				Else 
					PF8
					EMReadScreen last_page_check, 21, 24, 2
					If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
				END IF 
			Else 
				exit do
			END IF 
		LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"
		
		'If issuance_month = "08/16" then aug_total 	= HG_amt_issued
		'If issuance_month = "09/16" then sept_total = HG_amt_issued
		'If issuance_month = "10/16" then oct_total 	= HG_amt_issued
		'If issuance_month = "11/16" then nov_total 	= HG_amt_issued
		'If issuance_month = "12/16" then dec_total 	= HG_amt_issued
		If issuance_month = "01/17" then jan_total 		= HG_amt_issued
		If issuance_month = "02/17" then feb_total 		= HG_amt_issued
        If issuance_month = "03/17" then march_total 	= HG_amt_issued
		If issuance_month = "04/17" then april_total 	= HG_amt_issued
		If issuance_month = "05/17" then may_total 		= HG_amt_issued
		'msgbox issuance_month & vbcr & HG_amt_issued
		
		'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
		Do
			PF7
			EMReadScreen first_page_check, 20, 24, 2
		LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page
	NEXT
	
	'Adding client information to the array
	ReDim Preserve HG_array(6, 	case_count)	'This resizes the array 
	HG_array(case_number, 		case_count)		= MAXIS_case_number
	'HG_array (aug_2016, 		case_count) 	= aug_total
	'HG_array (sept_2016, 		case_count) 	= sept_total
	'HG_array (oct_2016, 		case_count) 	= oct_total
	'HG_array (nov_2016, 		case_count) 	= nov_total
	'HG_array (dec_2016, 		case_count) 	= dec_total
	HG_array (jan_2017, 		case_count) 	= jan_total
	HG_array (feb_2017, 		case_count) 	= feb_total
    HG_array (march_2017, 		case_count) 	= march_total
	HG_array (april_2017, 		case_count) 	= april_total
	HG_array (may_2017, 		case_count) 	= may_total
	
	STATS_counter = STATS_counter + 1		 'adds one instance to the stats counter
	MAXIS_case_number = ""
Next

Erase case_numbers_array
case_number_list = ""

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

msgbox "Excel is about to open. Close other Excel programs."
'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
'ObjExcel.Cells(1, 2).Value = "Aug 2016"					
'ObjExcel.Cells(1, 3).Value = "Sept 2016"
'ObjExcel.Cells(1, 4).Value = "Oct 2016"
'ObjExcel.Cells(1, 5).Value = "Nov 2016"
'ObjExcel.Cells(1, 6).Value = "Dec 2016"
ObjExcel.Cells(1, 2).Value = "Jan 2017"
ObjExcel.Cells(1, 3).Value = "Feb 2017"
ObjExcel.Cells(1, 4).Value = "March 2017"
ObjExcel.Cells(1, 5).Value = "April 2017"
ObjExcel.Cells(1, 6).Value = "May 2017"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2
'Updating the Excel spreadsheet based on what's happening in MAXIS----------------------------------------------------------------------------------------------------
For i = 0 to UBound(HG_array, 2)
	objExcel.cells(excel_row, 1).Value = HG_array(case_number, 	i)
	'objExcel.cells(excel_row, 2).Value = HG_array(aug_2016, 	i)
	'objExcel.cells(excel_row, 3).Value = HG_array(sept_2016, 	i)
	'objExcel.cells(excel_row, 4).Value = HG_array(oct_2016, 	i)
	'objExcel.cells(excel_row, 5).Value = HG_array(nov_2016, 	i)
	'objExcel.cells(excel_row, 6).Value = HG_array(dec_2016, 	i)
	objExcel.cells(excel_row, 2).Value = HG_array(jan_2017, 	i)
	objExcel.cells(excel_row, 3).Value = HG_array(feb_2017, 	i)
    objExcel.cells(excel_row, 4).Value = HG_array(march_2017, 	i)
	objExcel.cells(excel_row, 5).Value = HG_array(april_2017, 	i)
	objExcel.cells(excel_row, 6).Value = HG_array(may_2017, 	i)
	excel_row = excel_row + 1
Next 

col_to_use = 8
'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'Auto-fitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'script_end_procedure("Success! Please review the list generated.")