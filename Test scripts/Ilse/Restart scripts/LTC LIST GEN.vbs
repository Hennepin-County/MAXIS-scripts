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

'CONNECTS TO BlueZone
EMConnect ""

file_selection_path = "C:\Users\ilfe001\OneDrive - Hennepin County\Desktop\LTC GRH List Gen 11-2022.xlsx"


'The dialog is defined in the loop as it can change as buttons are pressed
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "Restart COLA Decimator at CASE/NOTE."
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a COLA Decimator list needs to be restared at the point of the Case noting portion."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
'dialog and dialog DO...Loop
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
'Select Excel row dialog
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to restart"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

excel_row_to_restart = "34887"

Do
	dialog Dialog1
	cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart

back_to_self 'This Do...loop gets back to SELF
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit
call navigate_to_MAXIS_screen("REPT", "ACTV")

do until ObjExcel.Cells(excel_row, 2).Value = "" 'shuts down when there's no more case numbers
	FACI_name = "" 'Resetting these variables
    vendor_number = ""
	FACI_type = ""
	GRH_DOC = ""
	AREP_name = ""
	DISA_waiver_type = ""
	SWKR_name = ""

	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	If trim(MAXIS_case_number) = "" then exit do
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
	    	ObjExcel.Cells(excel_row, 6).Value = trim(replace(FACI_type, "_", ""))
	    End if
	    'Navigates to DISA and checks the waiver type
	    EMWriteScreen "DISA", 20, 71
	    transmit
	    EMReadScreen DISA_waiver_type, 1, 14, 59
	    If DISA_waiver_type = "_" then DISA_waiver_type = ""
	    ObjExcel.Cells(excel_row, 7).Value = DISA_waiver_type
    End if
	excel_row = excel_row + 1 'setting up the script to check the next row.
loop

'formatting the cells
FOR i = 1 to 7
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your list has been created.")
