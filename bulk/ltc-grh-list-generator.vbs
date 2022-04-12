'Required for statistical purposes===============================================================================
name_of_script = "BULK - LTC-GRH LIST GENERATOR.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                      'manual run time in seconds
STATS_denomination = "C"       						 'C is for each CASE
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
call changelog_update("04/12/2022", "Added Vendor Number column to output of report.", "Ilse Ferris, Hennepin County")
call changelog_update("01/20/2017", "Added SWKR column. Updated BULK script to allow users to select what information is added (in addition to worker #, case #, client and FACI name).", "Ilse Ferris, Hennepin County")
call changelog_update("01/03/2017", "Added FACI type column. Reordered GRH DOC amt, waiver type and AREP columns.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(False)
'grabbing current footer month/year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'DIALOG TO DETERMINE WHERE TO GO IN MAXIS TO GET THE INFO
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 286, 130, "LTC-GRH list generator dialog"
  DropListBox 90, 10, 60, 15, "REPT/ACTV"+chr(9)+"REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  EditBox 235, 10, 20, 15, MAXIS_footer_month
  EditBox 260, 10, 20, 15, MAXIS_footer_year
  EditBox 90, 35, 190, 15, worker_number
  CheckBox 20, 90, 45, 10, "FACI type", FACI_type_checkbox
  CheckBox 70, 90, 60, 10, "GRH DOC amt", DOC_checkbox
  CheckBox 135, 90, 50, 10, "Waiver type", waiver_checkbox
  CheckBox 190, 90, 30, 10, "AREP", AREP_checkbox
  CheckBox 230, 90, 35, 10, "SWKR", SWKR_checkbox
  ButtonGroup ButtonPressed
    OkButton 175, 110, 50, 15
    CancelButton 230, 110, 50, 15
  Text 5, 55, 250, 10, "Enter7 digits of each worker number, (ex: x######), seperated by a comma."
  GroupBox 5, 75, 275, 30, "Select info to add (in addition to worker #, case #, client, FACI name and vendor #):"
  Text 170, 15, 65, 10, "Footer month/year:"
  Text 25, 40, 60, 10, "Worker number(s):"
  Text 35, 15, 55, 10, "Create list from:"
EndDialog

'DISPLAYS DIALOG
Do
	Do
		err_msg = ""
		Dialog Dialog1
		Cancel_without_confirmation
		If worker_number = "" then err_msg = err_msg & vbnewline & "* Enter at least one worker number."
		If isnumeric(MAXIS_footer_month) = false or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbnewline & "* Enter a valid 2-digit footer month."
		If isnumeric(MAXIS_footer_year) = false or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbnewline & "* Enter a valid 2-digit footer year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

'NAVIGATES BACK TO SELF TO FORCE THE FOOTER MONTH, THEN NAVIGATES TO THE SELECTED SCREEN
back_to_self
EMWriteScreen "________", 18, 43
call navigate_to_MAXIS_screen("rept", right(REPT_panel, 4))

'REVS can be in CM + 2 after the 16th of the month
If right(REPT_panel, 4) = "REVS" then
	EMWriteScreen MAXIS_footer_month, 20, 55
	EMWriteScreen MAXIS_footer_year, 20, 58
	transmit
End if

'CHECKS TO MAKE SURE WE'VE MOVED PAST SELF MENU. IF WE HAVEN'T, THE SCRIPT WILL STOP. AN ERROR MESSAGE SHOULD DISPLAY ON THE BOTTOM OF THE MENU.
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then script_end_procedure("Can't get past SELF menu. Check error message and try again!")

''Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

excel_row = 2
'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "MAXIS Case #"
ObjExcel.Cells(1, 3).Value = "Client Name"
ObjExcel.Cells(1, 4).Value = "Facility Name"
ObjExcel.Cells(1, 5).Value = "Vendor #"

col_to_use = 5 'Starting with 6 because cols 1-5 are already used

If FACI_type_checkbox = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "FACI Type"
	faci_type_col = col_to_use
	col_to_use = col_to_use + 1
End if

If DOC_checkbox = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "GRH DOC Amt"
	DOC_col = col_to_use
	col_to_use = col_to_use + 1
End if

If waiver_checkbox = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "Waiver type on first DISA panel found"
	waiver_col = col_to_use
	col_to_use = col_to_use + 1
End if

If AREP_checkbox = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "AREP"
	AREP_col = col_to_use
	col_to_use = col_to_use + 1
End if

If SWKR_checkbox = 1 then
	ObjExcel.Cells(1, col_to_use).Value = "SWKR"
	SWKR_col = col_to_use
	col_to_use = col_to_use + 1
End if

'formatting the cells
FOR i = 1 to col_to_use
	objExcel.Cells(1, i).Font.Bold = True		'bold font
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

'Splitting array for use by the for...next statement
worker_number_array = split(worker_number, ",")

For each worker in worker_number_array

	If trim(worker) = "" then exit for

	worker_ID = trim(worker)

	If REPT_panel = "REPT/ACTV" then 'THE REPT PANEL has all the information in different places than REVS and REVW
		worker_ID_col = 13
        case_number_col = 12
        client_name_col = 21
	Else
		worker_ID_col = 6
        case_number_col = 6
        client_name_col = 16
	End if

	EMReadScreen default_worker_number, 7, 21, worker_ID_col 'CHECKING THE CURRENT worker NUMBER. IF IT DOESN'T NEED TO CHANGE IT WON'T. OTHERWISE, THE SCRIPT WILL INPUT THE CORRECT NUMBER.
	If ucase(worker_ID) <> ucase(default_worker_number) then
		EMWriteScreen worker_ID, 21, worker_ID_col
		transmit
	End if

	'THIS DO...LOOP DUMPS THE CASE NUMBER AND NAME OF EACH CLIENT INTO A SPREADSHEET
	Do
		EMReadScreen last_page_check, 21, 24, 02
		row = 7 'defining the row to look at
		Do
			EMReadScreen MAXIS_case_number, 8, row, case_number_col  'grabbing case number
			EMReadScreen client_name, 15, row, client_name_col       ' grabbing client name
			If trim(MAXIS_case_number) = "" then exit do	         'quits if we're out of cases
			ObjExcel.Cells(excel_row, 1).Value = worker_ID
			ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
			ObjExcel.Cells(excel_row, 3).Value = trim(client_name)
			excel_row = excel_row + 1
			row = row + 1
			STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
		Loop until row = 19 or trim(MAXIS_case_number) = ""

		PF8 'going to the next screen
	Loop until last_page_check = "THIS IS THE LAST PAGE"
next

'NOW THE SCRIPT IS CHECKING STAT/FACI FOR EACH CASE.----------------------------------------------------------------------------------------------------

excel_row = 2 'Resetting the case row to investigate.
back_to_self 'This Do...loop gets back to SELF
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
transmit

do until ObjExcel.Cells(excel_row, 1).Value = "" 'shuts down when there's no more case numbers
	FACI_name = "" 'Resetting these variables
    vendor_number = ""
	FACI_type = ""
	GRH_DOC = ""
	AREP_name = ""
	DISA_waiver_type = ""
	SWKR_name = ""

	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	If MAXIS_case_number = "" then exit do
	back_to_SELF
    Call navigate_to_MAXIS_screen_review_PRIV("STAT", "FACI", is_this_priv)
    If is_this_priv = True then
        ObjExcel.Cells(excel_row, 3).Value = "Privileged Case." 'overwriting the client name to indicate priv case.
    Else
	    'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
	    Do
	    	EMReadScreen FACI_current_panel, 1, 2, 73
	    	EMReadScreen FACI_total_check, 1, 2, 78
	    	EMReadScreen in_year_check_01, 4, 14, 53
	    	EMReadScreen in_year_check_02, 4, 15, 53
	    	EMReadScreen in_year_check_03, 4, 16, 53
	    	EMReadScreen in_year_check_04, 4, 17, 53
	    	EMReadScreen in_year_check_05, 4, 18, 53
	    	EMReadScreen out_year_check_01, 4, 14, 77
	    	EMReadScreen out_year_check_02, 4, 15, 77
	    	EMReadScreen out_year_check_03, 4, 16, 77
	    	EMReadScreen out_year_check_04, 4, 17, 77
	    	EMReadScreen out_year_check_05, 4, 18, 77
	    	If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or _
	    	(in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
	    		currently_in_FACI = True
	    		exit do
	    	Elseif FACI_current_panel = FACI_total_check then
	    		currently_in_FACI = False
	    		exit do
	    	Else
	    		transmit
	    	End if
	    Loop until FACI_current_panel = FACI_total_check

	    'GETS FACI NAME AND PUTS IT IN SPREADSHEET, IF CLIENT IS IN FACI.
	    If currently_in_FACI = True then
	    	EMReadScreen FACI_name, 30, 6, 43
	    	EMReadScreen FACI_type, 2, 7, 43
            EmReadscreen vendor_number, 8, 5, 43
	    	'List of FACI types
	    	IF FACI_type = "41" then FACI_type = "41: NF-I"
	    	IF FACI_type = "42" then FACI_type = "42: NF-II"
	    	IF FACI_type = "43" then FACI_type = "43: ICF-DD"
	    	IF FACI_type = "44" then FACI_type = "44: Short stay in NF-I"
	    	IF FACI_type = "45" then FACI_type = "45: Short stay in NF-II"
	    	IF FACI_type = "46" then FACI_type = "46: Short stay in ICF-DD"
	    	IF FACI_type = "47" then FACI_type = "47: RTC - Not IMD"
	    	IF FACI_type = "48" then FACI_type = "48: Medical Hospital"
	    	IF FACI_type = "49" then FACI_type = "49: MSOP"
	    	IF FACI_type = "50" then FACI_type = "50: IMD/RTC"
	    	IF FACI_type = "51" then FACI_type = "51: Rule 31 CD_IMD"
	    	IF FACI_type = "52" then FACI_type = "52: Rule 36 MI-IMD"
	    	IF FACI_type = "53" then FACI_type = "53: IMD Hospitals"
	    	IF FACI_type = "55" then FACI_type = "55: Adult Foster Care/Rule 203"
	    	IF FACI_type = "56" then FACI_type = "56: GRH (Not FC or Rule 36)"
	    	IF FACI_type = "57" then FACI_type = "57: Rule 36 MI - Non-IMD"
	    	IF FACI_type = "60" then FACI_type = "60: Non-GRH"
	    	IF FACI_type = "61" then FACI_type = "61: Rule 31 CD - Non-IMD"
	    	IF FACI_type = "67" then FACI_type = "67: Family Violence Shelter"
	    	IF FACI_type = "68" then FACI_type = "68: County Correctional Facility"
	    	IF FACI_type = "69" then FACI_type = "69: Non-Cty Adult Correctional"

	    	EMReadScreen GRH_DOC, 8, 13, 45
	    	ObjExcel.Cells(excel_row, 4).Value = trim(replace(FACI_name, "_", ""))
            ObjExcel.Cells(excel_row, 5).Value = trim(replace(vendor_number, "_", ""))
	    	If FACI_type_checkbox = 1  	then ObjExcel.Cells(excel_row, faci_type_col).Value = trim(replace(FACI_type, "_", ""))
	    	If DOC_checkbox = 1 		then ObjExcel.Cells(excel_row, DOC_col).Value = trim(replace(GRH_DOC, "_", ""))
	    End if

	    IF AREP_checkbox = 1 then
	    	'NAVIGATES TO AREP, READS THE NAME, AND ADDS TO SPREADSHEET
	    	EMWriteScreen "AREP", 20, 71
	    	transmit
	    	EMReadScreen AREP_name, 37, 4, 32
	    	AREP_name = replace(AREP_name, "_", "")
	    	ObjExcel.Cells(excel_row, AREP_col).Value = AREP_name
	    END IF

	    If waiver_checkbox = 1 then
	    	'Navigates to DISA and checks the waiver type
	    	EMWriteScreen "DISA", 20, 71
	    	transmit
	    	EMReadScreen DISA_waiver_type, 1, 14, 59
	    	If DISA_waiver_type = "_" then DISA_waiver_type = ""
	    	ObjExcel.Cells(excel_row, waiver_col).Value = DISA_waiver_type
	    END IF

	    IF SWKR_checkbox = 1 then
	    	'NAVIGATES TO STAT/SWKR and reads the SWKR name
	    	EMWritescreen "SWKR", 20, 71
	    	transmit
	    	EMReadScreen SWKR_name, 34, 6, 32
	    	swkr_name = replace(swkr_name, "_", "")
	    	ObjExcel.Cells(excel_row, SWKR_col).Value = swkr_name
	    END IF
    End if
	excel_row = excel_row + 1 'setting up the script to check the next row.
loop

'formatting the cells
FOR i = 1 to col_to_use
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your list has been created.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/12/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/12/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/12/2022
'--All variables in dialog match mandatory fields-------------------------------04/12/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/12/2022--------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------04/12/2022--------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/12/2022--------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/12/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------04/12/2022--------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------04/12/2022
'--Out-of-County handling reviewed----------------------------------------------04/12/2022--------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/12/2022
'--BULK - review output of statistics and run time/count (if applicable)--------04/12/2022--------------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/12/2022
'--Incrementors reviewed (if necessary)-----------------------------------------04/12/2022
'--Denomination reviewed -------------------------------------------------------04/12/2022
'--Script name reviewed---------------------------------------------------------04/12/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------04/12/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------04/12/2022
'--comment Code-----------------------------------------------------------------04/12/2022
'--Update Changelog for release/update------------------------------------------04/12/2022
'--Remove testing message boxes-------------------------------------------------04/12/2022
'--Remove testing code/unnecessary code-----------------------------------------04/12/2022
'--Review/update SharePoint instructions----------------------------------------04/12/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/12/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------04/12/2022
'--Complete misc. documentation (if applicable)---------------------------------04/12/2022--------------------N/A
'--Update project team/issue contact (if applicable)----------------------------04/12/2022
