'Required for statistical purposes===============================================================================
name_of_script = "BULK - TIKL FROM LIST.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 37          'manual run time in seconds
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
call changelog_update("08/16/2021", "Added to ADMIN - MONTHLY TASKS MENU.", "Ilse Ferris, Hennepin County")
call changelog_update("06/27/2018", "Added/updated closing message.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'>>>>> Function to build dlg for manual entry <<<<<
FUNCTION build_manual_entry_dlg(case_number_array, TIKL_text)
	'Array for all case numbers
	'This was chosen over building a dlg with 50 variables
	REDim all_cases_array(50, 0)
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 331, 310, "Enter MAXIS case numbers"
		Text 10, 15, 140, 10, "Enter MAXIS case numbers below..."
		dlg_row = 30
		dlg_col = 10
		FOR i = 1 TO 50
			EditBox dlg_col, dlg_row, 55, 15, all_cases_array(i, 0)
			dlg_row = dlg_row + 20
			IF dlg_row = 230 THEN
				dlg_row = 30
				dlg_col = dlg_col + 65
			END IF
		NEXT
		Text 10, 240, 90, 10, "Enter your TIKL text..."
		EditBox 10, 255, 310, 15, TIKL_text
		Text 10, 280, 80, 10, "TIKL Date (MM/DD/YY):"
		EditBox 95, 275, 80, 15, TIKL_date
		ButtonGroup ButtonPressed
			OkButton 220, 290, 50, 15
			CancelButton 270, 290, 50, 15
	EndDialog

	'Calling the dlg within the function
	DO
		'err_msg handling
		err_msg = ""
		DIALOG Dialog1
			cancel_confirmation
			FOR i = 1 TO 50
				all_cases_array(i, 0) = replace(all_cases_array(i, 0), " ", "")
				IF all_cases_array(i, 0) <> "" THEN
					IF len(all_cases_array(i, 0)) > 8 THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " is too long to be a valid MAXIS case number."
					IF isnumeric(all_cases_array(i, 0)) = FALSE THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " contains alphabetic characters. These are not valid."
				END IF
			NEXT
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""

	'building the array
	case_number_array = ""
	FOR i = 1 TO 50
		IF all_cases_array(i, 0) <> "" THEN case_number_array = case_number_array & all_cases_array(i, 0) & "~~~"
	NEXT
END FUNCTION

'----------FUNCTIONS----------
'-----This function needs to be added to the FUNCTIONS FILE-----
'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

'The script===========================
EMConnect ""
CALL check_for_MAXIS(true)

'>>>>> loading the main dialog <<<<<
 Dialog1 = ""
   BeginDialog Dialog1, 0, 0, 201, 65, "TIKL from List"
   DropListBox 5, 40, 80, 10, "Manual Entry"+chr(9)+"REPT/ACTV"+chr(9)+"Excel File", run_mode
   ButtonGroup ButtonPressed
   OkButton 90, 40, 50, 15
   CancelButton 140, 40, 50, 15
   Text 10, 10, 185, 25, "Please select a run mode for the script. You can either enter the case numbers manually, from REPT/ACTV, or from an Excel file..."
EndDialog
   
DIALOG Dialog1
cancel_without_confirmation
'>>>>> the script has different ways of building case_number_array
IF run_mode = "Manual Entry" THEN
	CALL build_manual_entry_dlg(case_number_array, TIKL_text)

ELSEIF run_mode = "REPT/ACTV" THEN
	'script_end_procedure("This mode is not yet supported.")
	CALL find_variable("User: ", worker_number, 7)
    
    '>>>>> THE DLG for REPT/ACTV mode<<<<<
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 231, 110, "Enter worker number and TIKL text..."
      EditBox 145, 10, 65, 15, worker_number
      EditBox 10, 50, 215, 15, TIKL_text
      EditBox 90, 70, 80, 15, TIKL_date
      ButtonGroup ButtonPressed
        OkButton 65, 90, 50, 15
        CancelButton 120, 90, 50, 15
      Text 10, 15, 130, 10, "Please enter the 7-digit worker number:"
      Text 10, 35, 95, 10, "Enter your TIKL text..."
      Text 10, 75, 80, 10, "TIKL Date (MM/DD/YY):"
    EndDialog
    Do 
	    Do
	    	err_msg = ""
	    	DIALOG Dialog1
	    	cancel_confirmation
	    	worker_number = trim(worker_number)
	    	IF trim(worker_number) = "" or len(worker_number) <> 7 THEN err_msg = err_msg & vbCr & "* You must enter the full 7-digit worker number."
	    	If trim(TIKL_text) = "" then err_msg = err_msg & vbCr & "* Enter the TIKL text."
            If trim(TIKL_date) = "" or isDate(TIKL_date) = false then err_msg = err_msg & vbCr & "* You must enter a valid TIKL date."
	    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	    LOOP UNTIL err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

	CALL check_for_MAXIS(false)

	'Checking that case number is blank so as to get a full REPT/ACTV
	CALL find_variable("Case Nbr: ", MAXIS_case_number, 8)
	MAXIS_case_number = replace(MAXIS_case_number, "_", " ")
	MAXIS_case_number = trim(MAXIS_case_number)
	IF MAXIS_case_number <> "" THEN
		back_to_SELF
		EMWriteScreen "________", 18, 43
	END IF
	'Checking that MAXIS is not already in REPT/ACTV so as to get a full REPT/ACTV
	EMReadScreen at_REPT_ACTV, 4, 2, 48
	IF at_REPT_ACTV = "ACTV" THEN back_to_SELF

	CALL navigate_to_MAXIS_screen("REPT", "ACTV")
	CALL write_value_and_transmit(worker_number, 21, 13)
	'Making sure we are at the beginning of REPT/ACTV
	DO
		PF7
		EMReadScreen page_one, 2, 3, 78
		IF isnumeric(page_one) = false then page_one = page_one * 1 'this is future proofing since reading variables keep switching back from numeric and non numeric.
	LOOP UNTIL page_one = 1

	rept_row = 7
	DO
		last_page_check = ""
		EMReadScreen MAXIS_case_number, 8, rept_row, 12
		MAXIS_case_number = trim(MAXIS_case_number)
		IF MAXIS_case_number <> "" THEN
			case_number_array = case_number_array & MAXIS_case_number & "~~~"
			rept_row = rept_row + 1
			IF rept_row = 19 THEN
				rept_row = 7
				PF8
				EMReadScreen last_page_check, 4, 24, 14			'this prevents the script from erroring out if the worker only has one completely full page of cases.
				If last_page_check = "LAST" THEN EXIT DO
			END IF
		ELSE
			EXIT DO
		END IF
	LOOP

ELSEIF run_mode = "Excel File" THEN
	'Opening the Excel file
	DO
		call file_selection_system_dialog(excel_file_path, ".xlsx")	'Selects an excel file, adds it to excel_file_path

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

	'Gathering the information from the user about the fields in Excel to look for.
    '>>>>>DLG for Excel mode<<<<<
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 256, 135, "TIKL Information"
      EditBox 220, 10, 25, 15, excel_col
      EditBox 65, 30, 40, 15, excel_row
      EditBox 190, 30, 40, 15, end_row
      EditBox 10, 70, 235, 15, TIKL_text
      EditBox 95, 90, 80, 15, TIKL_date
      ButtonGroup ButtonPressed
        OkButton 130, 115, 55, 15
        CancelButton 190, 115, 60, 15
      Text 10, 15, 205, 10, "Please enter the column containing the MAXIS case numbers..."
      Text 10, 35, 50, 10, "Row to start..."
      Text 135, 35, 50, 10, "Row to end..."
      Text 10, 55, 230, 10, "Please enter your TIKL text. Separate new lines with semi-colons..."
      Text 10, 95, 80, 10, "TIKL Date (MM/DD/YY):"
    EndDialog
    
    Do 
	    DO
	    	err_msg = ""
	    	DIALOG Dialog1
	    		cancel_confirmation 
	    		IF isnumeric(excel_col) = FALSE AND len(excel_col) > 2 THEN
	    			err_msg = err_msg & vbCr & "* Please do not use such a large column. The script cannot handle it."
	    		ELSE
	    			IF (isnumeric(right(excel_col, 1)) = TRUE AND isnumeric(left(excel_col, 1)) = FALSE) OR (isnumeric(right(excel_col, 1)) = FALSE AND isnumeric(left(excel_col, 1)) = TRUE) THEN
	    				err_msg = err_msg & vbCr & "* Please use a valid Column indicator. " & excel_col & " contains BOTH a letter and a number."
	    			ELSE
	    				call convert_excel_letter_to_excel_number(excel_col)
	    				IF isnumeric(excel_row) = false or isnumeric(end_row) = false THEN err_msg = err_msg & vbCr & "* Please enter the Excel rows as numeric characters."
	    				IF end_row = "" THEN err_msg = err_msg & vbCr & "* Please enter an end to the search. The script needs to know when to stop searching."
	    			END IF
	    		END IF
	    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	    LOOP UNTIL err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

	CALL check_for_MAXIS(false)
	'Generating a TIKL for each case.
	FOR i = excel_row TO end_row
		IF objExcel.Cells(i, excel_col).Value <> "" THEN
			case_number_array = case_number_array & objExcel.Cells(i, excel_col).Value & "~~~"
		END IF
	NEXT
END IF

'The business of sending TIKLSs
case_number_array = trim(case_number_array)
case_number_array = split(case_number_array, "~~~")

privileged_array = ""

FOR EACH MAXIS_case_number IN case_number_array
	IF MAXIS_case_number <> "" THEN
        'Checking PRIV status 
        Call navigate_to_MAXIS_screen_review_PRIV("DAIL", "WRIT", is_this_priv)
		If is_this_priv = True then 
			privileged_array = privileged_array & MAXIS_case_number & "~~~"
		ELSE
			Call create_TIKL(TIKL_text, 0, TIKL_date, False, TIKL_note_text)
			STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
		END IF
	END IF
NEXT

IF privileged_array <> "" THEN
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a TIKL for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your list has been completed.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation 
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/16/2021
'--Tab orders reviewed & confirmed----------------------------------------------08/16/2021  
'--Mandatory fields all present & Reviewed--------------------------------------08/16/2021
'--All variables in dialog match mandatory fields-------------------------------08/16/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------08/16/2021----------------N/A: Only sets TIKL
'--CASE:NOTE Header doesn't look funky------------------------------------------08/16/2021----------------N/A: Only sets TIKL
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------08/16/2021----------------N/A: Only sets TIKL
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/16/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------08/16/2021----------------N/A
'--PRIV Case handling reviewed -------------------------------------------------08/16/2021
'--Out-of-County handling reviewed----------------------------------------------08/16/2021----------------N/A: TIKL's can be set out-of-coumty
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/16/2021
'--BULK - review output of statistics and run time/count (if applicable)--------08/16/2021----------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/16/2021
'--Incrementors reviewed (if necessary)-----------------------------------------08/16/2021
'--Denomination reviewed -------------------------------------------------------08/16/2021
'--Script name reviewed---------------------------------------------------------08/16/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/16/2021

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------08/16/2021
'--comment Code-----------------------------------------------------------------08/16/2021
'--Update Changelog for release/update------------------------------------------08/16/2021
'--Remove testing message boxes-------------------------------------------------08/16/2021
'--Remove testing code/unnecessary code-----------------------------------------08/16/2021
'--Review/update SharePoint instructions----------------------------------------08/16/2021
'--Review Best Practices using BZS page ----------------------------------------08/16/2021---------------N/A
'--Review script information on SharePoint BZ Script List-----------------------08/16/2021---------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------08/16/2021---------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------08/16/2021
'--Complete misc. documentation (if applicable)---------------------------------08/16/2021
'--Update project team/issue contact (if applicable)----------------------------08/16/2021