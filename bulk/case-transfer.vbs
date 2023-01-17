'Script Created by Casey Love from Ramsey County
'Required for statistical purposes==========================================================================================
name_of_script = "BULK - CASE TRANSFER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 20                	'manual run time in seconds
STATS_denomination = "I"       			'I is for each Item
'END OF stats block=========================================================================================================

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
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("07/02/2021", "BUG on BULK Case Transfer that was pulling too many pending cases into the list. Fixed now.", "Casey Love, Hennepin County")
CALL changelog_update("04/16/2020", "Updated BULK Case Transfer to accomodate pending cases better.", "Casey Love, Hennepin County")
CALL changelog_update("10/04/2019", "Added functionality to read cases from PND2 as well.##~## ##~## As this is new functionality and testing ability is limited, please report any issues are errors to the BlueZone Script Team.##~##", "Casey Love, Hennepin County")
CALL changelog_update("01/24/2019", "BUG fixed that caused an error if there are no MFIP eligibility results.", "Casey Love, Hennepin County")
CALL changelog_update("01/12/2018", "Entering a supervisor X-Number in the Workers to Check will pull all X-Numbers listed under that supervisor in MAXIS. Addiional bug fix where script was missing cases.", "Casey Love, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code
'Connects to BlueZone
EMConnect ""

'Shows dialogDialog pull_rept_data_into_Excel_dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 781, 240, "Select Parameters for Cases to Transfer"
  EditBox 75, 20, 165, 15, worker_number
  CheckBox 450, 20, 240, 10, "Check here to have the script query all the cases listed on REPT/ACTV.", query_all_actv_check
  CheckBox 80, 90, 40, 10, "SNAP", SNAP_check
  CheckBox 80, 105, 90, 10, "Exclude all SNAP cases", exclude_snap_check
  CheckBox 80, 115, 75, 10, "SNAP ONLY cases", SNAP_Only_check
  CheckBox 85, 145, 60, 10, "ABAWD cases", SNAP_ABAWD_check
  CheckBox 85, 155, 90, 10, "Uncle Harry SNAP", SNAP_UH_check
  CheckBox 185, 90, 75, 10, "Any/All Cash", all_cash_check
  CheckBox 185, 105, 25, 10, "GA", ga_check
  CheckBox 230, 105, 30, 10, "MSA", msa_check
  CheckBox 185, 115, 95, 10, "Exclude all GA/MSA cases", exclude_ga_msa_check
  CheckBox 185, 130, 25, 10, "RCA", rca_check
  CheckBox 185, 140, 90, 10, "Exclude all RCA cases", exclude_RCA_check
  CheckBox 295, 105, 30, 10, "MFIP", mfip_check
  CheckBox 340, 105, 30, 10, "DWP", DWP_check
  CheckBox 295, 115, 95, 10, "Exclude all MFIP/DWP", exclude_mfip_dwp_check
  CheckBox 295, 125, 70, 10, "MFIP ONLY cases", MFIP_Only_check
  CheckBox 295, 150, 85, 10, "Child Only MFIP cases", child_only_mfip_check
  CheckBox 295, 160, 90, 10, "Only monthly reporters", mont_rept_check
  CheckBox 295, 170, 90, 10, "MFIP cases with at least", MFIP_tanf_check
  EditBox 385, 165, 20, 15, tanf_months
  CheckBox 400, 90, 50, 10, "Child Care", ccap_check
  CheckBox 400, 105, 95, 10, "Exclude Child Care cases", exclude_ccap_check
  CheckBox 505, 90, 40, 10, "GRH", GRH_check
  CheckBox 505, 105, 75, 10, "Exclude GRH cases", exclude_grh_check
  CheckBox 505, 115, 75, 10, "GRH ONLY cases", GRH_Only_check
  CheckBox 600, 90, 40, 10, "HC", HC_check
  CheckBox 600, 105, 75, 10, "Exclude HC cases", exclude_HC_check
  CheckBox 600, 115, 75, 10, "HC ONLY cases", HC_Only_check
  CheckBox 605, 140, 50, 10, "MSP Active", HC_msp_check
  CheckBox 605, 150, 40, 10, "Adult MA", adult_hc_check
  CheckBox 605, 160, 45, 10, "Family MA", family_hc_check
  CheckBox 605, 170, 50, 10, "Waiver HC", waiver_HC_check
  CheckBox 605, 180, 40, 10, "LTC HC", ltc_HC_check
  CheckBox 680, 90, 75, 10, "EA/EGA Pending", EA_check
  CheckBox 680, 105, 95, 10, "Exclude EA/EGA Pending", exclude_ea_check
  CheckBox 10, 195, 170, 10, "Exclude all cases with any Pending Program", exclude_pending_check
  CheckBox 10, 205, 170, 10, "Include ONLY cases with Pending Programs", include_only_pending_check
  CheckBox 10, 215, 120, 10, "Exclude all cases with IEVS DAILs", exclude_ievs_check
  CheckBox 10, 225, 120, 10, "Exclude all cases with PARIS DAILs", exclude_paris_check
  EditBox 750, 195, 25, 15, case_found_limit
  ButtonGroup ButtonPressed
    OkButton 670, 220, 50, 15
    CancelButton 725, 220, 50, 15
  GroupBox 290, 140, 190, 45, "MFIP Details"
  GroupBox 600, 130, 60, 65, "HC Details"
  Text 290, 65, 315, 20, "Select the criteria you want the script to sort by. The script will then generate an Excel Spreadsheet of all the cases in the indicated worker number(s) that meet your selected criteria."
  Text 240, 210, 330, 30, "After the script searches all of the Workers to Check to find the cases that meet the inidcated criteria, it will have a complete spreadsheet to review. DO NOT MAKE CHANGES TO THIS SPREADSHEET. Another dialog will allow you enter the information for transfer before any transfers start."
  Text 410, 170, 65, 10, "TANF Months used."
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 455, 30, 235, 10, " This will not give a transfer option but will look for all possible criteria."
  Text 250, 20, 180, 20, "Note: please enter the entire 7-digit number x1 number (x100abc), separated by a comma."
  Text 330, 50, 220, 10, "Check all that apply - What type of cases do you want to transfer?"
  Text 250, 5, 100, 10, "***Case Parameters to Pull***"
  Text 615, 200, 130, 10, "Limit the number of cases available to:"
  GroupBox 80, 130, 80, 40, "SNAP Details"
  Text 5, 90, 75, 10, "Programs to Include:"
EndDialog
Do
	Do
		Dialog Dialog1
		cancel_without_confirmation
		err_msg = ""
		IF worker_number = "" then err_msg = err_msg & vbCr & "You must Select an X-Number to pull cases from."
		IF query_all_actv_check = unchecked AND snap_check = unchecked AND mfip_check = unchecked AND DWP_check = unchecked AND EA_check = unchecked AND HC_check = unchecked AND ga_check = unchecked AND msa_check = unchecked AND GRH_check = unchecked Then err_msg = err_msg & _
		  vbCr & "You must select a type of program for the cases you want to transfer, please pick one - SNAP, MFIP, DWP, EA, HC, GA, MSA, or GRH"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop until err_msg = ""
	err_msg = ""
	If SNAP_check = unchecked then
		IF SNAP_Only_check = checked OR SNAP_ABAWD_check = checked OR SNAP_UH_check = checked then err_msg = err_msg & vbCr & "If you select SNAP details, you must filter FOR SNAP cases - Check the SNAP box"
	End If
	IF mfip_check = unchecked then
		IF MFIP_Only_check = checked OR MFIP_tanf_check = checked OR child_only_mfip_check = checked OR mont_rept_check = checked then err_msg = err_msg & vbCr & " If you select MFIP details, you must filter FOR MFIP cases - check the MFIP box"
	End If
	If MFIP_tanf_check = checked AND tanf_months = "" then err_msg = err_msg & vbCr & "If you want to filter for a certain number of TANF months, you must indicate how many months you want"
	IF HC_check = unchecked then
		If HC_msp_check = checked OR adult_hc_check = checked OR family_hc_check = checked OR ltc_HC_check = checked OR waiver_HC_check = checked then err_msg = err_msg & vbCr & "If you select HC details, you must filter FOR HC cases - check the HC Box"
	End If
	IF snap_check = checked AND exclude_snap_check = checked then err_msg = err_msg & vbCr & "You cannot filter for SNAP and Exclude SNAP cases - please pick one"
	IF mfip_check = checked AND exclude_mfip_dwp_check = checked then err_msg = err_msg & vbCr & "You cannot filter for MFIP and Exclude MFIP cases - please pick one"
	IF ccap_check = checked AND exclude_ccap_check = checked then err_msg = err_msg & vbCr & "You cannot filter for Child Care and Exclude Child Care - please pick one"
	IF EA_check = checked AND exclude_ea_check = checked then err_msg = err_msg & vbCr & "You cannot filter for EA/EGA and Exclude EA/EGA cases - please pick one"
	IF HC_check = checked AND exclude_HC_check = checked then err_msg = err_msg & vbCr & "You cannot filter for HC and Exclude HC cases - please pick one"
	If exclude_ga_msa_check = checked then
		IF ga_check = checked OR msa_check = checked then err_msg = err_msg & vbCr & "You cannot filter for GA and/or MSA and Exclude GA/MSA cases - please pick one"
	End If
	If GRH_check = checked AND exclude_grh_check = checked then err_msg = err_msg & vbCr & "You cannot filter for GRH and Exclude GRH cases - please pick one"
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
Loop until err_msg = ""

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(False)

'In order to make the code a little cleaner, this sets all the option check boxes to checked when the Query All option exists.
IF query_all_actv_check = checked THEN
	SNAP_ABAWD_check = checked
	SNAP_UH_check = checked
	MFIP_tanf_check = checked
	child_only_mfip_check = checked
	mont_rept_check = checked
	HC_msp_check = checked
	adult_hc_check = checked
	family_hc_check = checked
	ltc_HC_check = checked
	waiver_HC_check = checked
	exclude_ievs_check = checked
	exclude_paris_check = checked
	ccap_check = checked
End IF

If case_found_limit <> "" Then case_found_limit = abs(case_found_limit)

MsgBox "The script will now create an Excel Spreadsheet to display case information" & _
   vbCr & "Please do not alter this spreadsheet until after the script has completed." & _
   vbCR & "Altering the spreadsheet will not change how the script runs and will only cause the data listed on it to be incorrect."


'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "NEXT REVW"
objExcel.Cells(1, 4).Font.Bold = TRUE

STATS_counter = STATS_counter + 1

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 5 because cols 1-4 are already used

'Setting the Program specific Excel col - the program headings will always populate but the more specific options will only populate if requested
ObjExcel.Cells(1, col_to_use).Value = "SNAP"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
snap_actv_col = col_to_use
col_to_use = col_to_use + 1
SNAP_letter_col = convert_digit_to_excel_column(snap_actv_col)

IF SNAP_ABAWD_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "ABAWD?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	ABAWD_actv_col = col_to_use
	col_to_use = col_to_use + 1
	ABAWD_letter_col = convert_digit_to_excel_column(ABAWD_actv_col)
End if

IF SNAP_UH_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "Unc Har?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	UH_actv_col = col_to_use
	col_to_use = col_to_use + 1
	UH_letter_col = convert_digit_to_excel_column(UH_actv_col)
End if

ObjExcel.Cells(1, col_to_use).Value = "Cash 1"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
cash_one_prog_col = col_to_use
col_to_use = col_to_use + 1
cash_one_prog_letter_col = convert_digit_to_excel_column(cash_one_prog_col)

ObjExcel.Cells(1, col_to_use).Value = "Status"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
cash_one_actv_col = col_to_use
col_to_use = col_to_use + 1
cash_one_letter_col = convert_digit_to_excel_column(cash_one_actv_col)

ObjExcel.Cells(1, col_to_use).Value = "Cash 2"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
cash_two_prog_col = col_to_use
col_to_use = col_to_use + 1
cash_two_prog_letter_col = convert_digit_to_excel_column(cash_two_prog_col)

ObjExcel.Cells(1, col_to_use).Value = "Status"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
cash_two_actv_col = col_to_use
col_to_use = col_to_use + 1
cash_two_letter_col = convert_digit_to_excel_column(cash_two_actv_col)

If MFIP_tanf_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "TANF"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	TANF_mo_col = col_to_use
	col_to_use = col_to_use + 1
	TANF_letter_col = convert_digit_to_excel_column(TANF_mo_col)
End if

If child_only_mfip_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "Child Only?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	child_only_col = col_to_use
	col_to_use = col_to_use + 1
	Child_letter_col = convert_digit_to_excel_column(child_only_col)
End if

If mont_rept_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HRF?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	mont_rept_col = col_to_use
	col_to_use = col_to_use + 1
	mont_rept_letter_col = convert_digit_to_excel_column(mont_rept_col)
End if

IF ccap_check = checked OR exclude_ccap_check = checked THEN
	ObjExcel.Cells(1, col_to_use).Value = "CCAP"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	ccap_col = col_to_use
	col_to_use = col_to_use + 1
	ccap_letter_col = convert_digit_to_excel_column(ccap_col)
End If

ObjExcel.Cells(1, col_to_use).Value = "HC"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
hc_actv_col = col_to_use
col_to_use = col_to_use + 1
hc_letter_col = convert_digit_to_excel_column(hc_actv_col)

If HC_msp_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "MSP"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	MSP_actv_col = col_to_use
	col_to_use = col_to_use + 1
	MSP_letter_col = convert_digit_to_excel_column(MSP_actv_col)
End if

If adult_hc_check = checked OR family_hc_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HC Type"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_type_col = col_to_use
	col_to_use = col_to_use + 1
	HC_type_letter_col = convert_digit_to_excel_column(HC_type_col)
End if

If ltc_HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "LTC?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	LTC_col = col_to_use
	col_to_use = col_to_use + 1
	LTC_letter_col = convert_digit_to_excel_column(LTC_col)
End if

If waiver_HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "Waiver?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	Waiver_col = col_to_use
	col_to_use = col_to_use + 1
	Waiver_letter_col = convert_digit_to_excel_column(Waiver_col)
End if

ObjExcel.Cells(1, col_to_use).Value = "EA/EGA"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
EA_actv_col = col_to_use
col_to_use = col_to_use + 1
EA_letter_col = convert_digit_to_excel_column(EA_actv_col)

ObjExcel.Cells(1, col_to_use).Value = "GRH"
objExcel.Cells(1, col_to_use).Font.Bold = TRUE
GRH_actv_col = col_to_use
col_to_use = col_to_use + 1
GRH_letter_col = convert_digit_to_excel_column(GRH_actv_col)

If exclude_ievs_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "IEVS?"
	ObjExcel.Cells(1, col_to_use).Font.Bold = True
	ievs_col = col_to_use
	col_to_use = col_to_use + 1
	ievs_letter_col = convert_digit_to_excel_column(ievs_col)
End If

If exclude_paris_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "PARIS?"
	ObjExcel.Cells(1, col_to_use).Font.Bold = True
	paris_col = col_to_use
	col_to_use = col_to_use + 1
	paris_letter_col = convert_digit_to_excel_column(paris_col)
End If

IF query_all_actv_check = unchecked THEN
	ObjExcel.Cells(1, col_to_use).Value = "Transferred?"
	ObjExcel.Cells(1, col_to_use).Font.Bold = TRUE
	xfered_col = col_to_use
	col_to_use = col_to_use + 1
	xfered_letter_col = convert_digit_to_excel_column(xfered_col)
End If

IF query_all_actv_check = unchecked THEN
	ObjExcel.Cells(1, col_to_use).Value = "New Worker"
	ObjExcel.Cells(1, col_to_use).Font.Bold = TRUE
	new_worker_col = col_to_use
	col_to_use = col_to_use + 1
	new_worker_letter_col = convert_digit_to_excel_column(new_worker_col)
End IF

'If worker numbers are litsted - this will create an array of workers to check
x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

'formatting array
For each x1_number in x1s_from_dialog
	x1_number = trim(ucase(x1_number))					'Formatting the x numbers so there are no errors
	Call navigate_to_MAXIS_screen ("REPT", "USER")		'This part will check to see if the x number entered is a supervisor of anyone
	PF5
	PF5
	EMWriteScreen x1_number, 21, 12
	transmit
	EMReadScreen sup_id_check, 7, 7, 5					'This is the spot where the first person is listed under this supervisor
	IF sup_id_check <> "       " Then 					'If this frist one is not blank then this person is a supervisor
		supervisor_array = trim(supervisor_array & " " & x1_number)		'The script will add this x number to a list of supervisors
	Else
		If worker_array = "" then						'Otherwise this x number is added to a list of workers to run the script on
			worker_array = trim(x1_number)
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	End If
	PF3
Next

If supervisor_array <> "" Then 				'If there are any x numbers identified as a supervisor, the script will run the function above
	Call create_array_of_all_active_x_numbers_by_supervisor (more_workers_array, supervisor_array)
	workers_to_add = join(more_workers_array, ", ")
	If worker_array = "" then				'Adding all x numbers listed under the supervisor to the worker array
		worker_array = workers_to_add
	Else
		worker_array = worker_array & ", " & trim(ucase(workers_to_add))
	End if
End If

'Split worker_array
worker_array = split(worker_array, ", ")

'Arrays that need delcaring and resizing
Dim All_case_information_array ()
Dim Full_case_list_array()
ReDim All_case_information_array (24, 0)
ReDim Full_case_list_array(12,0)
Dim eligible_members_array ()
Dim non_mfip_members_array ()
Dim SNAP_HH_Array ()

'Setting the variable for what's to come
excel_row = 2
all_case_numbers_array = "*"
m = 0

'Script starts by collecting a list of all the cases and the programs as listed on REPT/ACTV
'This information is added to the first array called - Full_case_list_array. The values of this array are as follows:
'(0,#) - Case Number
'(1,#) - Client Name
'(2,#) = Review Date
'(3,#) = Cash 1 Type
'(4,#) = Cash 1 Status
'(5,#) = Cash 2 Type
'(6,#) = Cash 2 Status
'(7,#) = SNAP Status
'(8,#) = HC Status
'(9,#) = EA Status
'(10,#) = GRH Status
'(11,#) = Worker's X Number
'(12,#) = CCAP Status

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
    If include_only_pending_check = unchecked Then
    	Call navigate_to_MAXIS_screen("rept", "actv")
    	EMWriteScreen worker, 21, 13
    	transmit
    	EMReadScreen user_worker, 7, 21, 71		'
    	EMReadScreen p_worker, 7, 21, 13
    	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

    	PF5 'Changes to case number sort for a better variety of cases.
    	'Skips workers with no info
    	EMReadScreen has_content_check, 1, 7, 8
    	If has_content_check <> " " then
    		'Grabbing each case number on screen
    		Do
    			'Set variable for next do...loop
    			MAXIS_row = 7

    			'Checking for the last page of cases.
    			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent

    			Do
                    Redim Preserve Full_case_list_array (Ubound(Full_case_list_array,1), m)

    				cash_one_type = ""
    				cash_two_type = ""

    				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12		'Reading case number
    				EMReadScreen client_name, 21, MAXIS_row, 21		'Reading client name
    				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
    				EMReadScreen cash_one_status, 1, MAXIS_row, 54		'Reading cash status
    					IF cash_one_status = "A" or cash_one_status = "P" then EMReadScreen cash_one_type, 2, MAXIS_row, 51
    				EMReadScreen cash_two_status, 1, MAXIS_row, 59
    					IF cash_two_status = "A" or cash_two_status = "P" then EMReadScreen cash_two_type, 2, MAXIS_row, 56
    				EMReadScreen SNAP_status, 1, MAXIS_row, 61		'Reading SNAP status
    				EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status
    				EMReadScreen EA_status, 1, MAXIS_row, 67			'Reading EA status
    				EMReadScreen GRH_status, 1, MAXIS_row, 70			'Reading GRH status
    				EMReadScreen CCAP_status, 1, MAXIS_row, 80 			'Reading CCAP Status

    				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end
                    next_revw_date = trim(next_revw_date)

    				Full_case_list_array(0,m) = MAXIS_case_number
    				Full_case_list_array(1,m) = client_name
    				Full_case_list_array(2,m) = replace(next_revw_date, " ", "/")
    				Full_case_list_array(3,m) = cash_one_type
    				Full_case_list_array(4,m) = cash_one_status
    				Full_case_list_array(5,m) = cash_two_type
    				Full_case_list_array(6,m) = cash_two_status
    				Full_case_list_array(7,m) = SNAP_status
    				Full_case_list_array(8,m) = HC_status
    				Full_case_list_array(9,m) = EA_status
    				Full_case_list_array(10,m) = GRH_status
    				Full_case_list_array(11,m) = worker
    				Full_case_list_array(12,m) = CCAP_status

    				' Redim Preserve Full_case_list_array (Ubound(Full_case_list_array,1), Ubound(Full_case_list_array,2)+1) 'Resize the array for the next case

    				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
    				MAXIS_case_number = trim(MAXIS_case_number)
    				If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
    				all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

    				MAXIS_row = MAXIS_row + 1
    				MAXIS_case_number = ""			'Blanking out variable
    				m = m + 1
    			Loop until MAXIS_row = 19
    			PF8
    		Loop until last_page_check = "THIS IS THE LAST PAGE"
    	End if
    End If

    back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
    Call navigate_to_MAXIS_screen("rept", "pnd2")
    EMWriteScreen worker, 21, 13
    transmit

    row = 1
    col = 1
    EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
    If row <> 0 Then transmit

    EMReadScreen user_worker, 7, 21, 71		'
    EMReadScreen p_worker, 7, 21, 13
    IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

    PF5 'Changes to case number sort for a better variety of cases.
    row = 1
    col = 1
    EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
    If row <> 0 Then transmit
    'Skips workers with no info
    EMReadScreen has_content_check, 1, 7, 12
    If has_content_check <> " " then
        'Grabbing each case number on screen
        Do
            'Set variable for next do...loop
            MAXIS_row = 7

            Do
                EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5		'Reading case number

                'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
                MAXIS_case_number = trim(MAXIS_case_number)
                If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") = 0 then
                    all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

                    Redim Preserve Full_case_list_array (Ubound(Full_case_list_array,1), m)

                    cash_one_type = ""
                    cash_two_type = ""

                    EMReadScreen client_name, 21, MAXIS_row, 16		'Reading client name
                    EMReadScreen cash_status, 1, MAXIS_row, 54		'Reading cash status
                        IF cash_status = "A" or cash_status = "P" then EMReadScreen cash_type, 2, MAXIS_row, 56
                    EMReadScreen SNAP_status, 1, MAXIS_row, 62		'Reading SNAP status
                    EMReadScreen HC_status, 1, MAXIS_row, 65			'Reading HC status
                    EMReadScreen EA_status, 1, MAXIS_row, 68			'Reading EA status
                    EMReadScreen GRH_status, 1, MAXIS_row, 72			'Reading GRH status
                    EMReadScreen CCAP_status, 1, MAXIS_row, 80 			'Reading CCAP Status

                    EMReadScreen next_line, 14, MAXIS_row + 1, 17
                    If next_line = "ADDITIONAL APP" Then
                        MAXIS_row = MAXIS_row + 1
                        EMReadScreen cash_add_status, 1, MAXIS_row, 54		'Reading cash status
                            IF cash_add_status = "A" or cash_add_status = "P" then EMReadScreen cash_add_type, 2, MAXIS_row, 56
                        EMReadScreen SNAP_add_status, 1, MAXIS_row, 62		'Reading SNAP status
                        EMReadScreen HC_add_status, 1, MAXIS_row, 65			'Reading HC status
                        EMReadScreen EA_add_status, 1, MAXIS_row, 68			'Reading EA status
                        EMReadScreen GRH_add_status, 1, MAXIS_row, 72			'Reading GRH status
                        EMReadScreen CCAP_add_status, 1, MAXIS_row, 80 			'Reading CCAP Status
                    End If

                    Full_case_list_array(0,m) = MAXIS_case_number
                    Full_case_list_array(1,m) = client_name
                    If cash_status = "A" OR cash_status = "P" Then Full_case_list_array(3,m) = cash_type
                    If cash_status = "A" OR cash_status = "P" Then Full_case_list_array(4,m) = "P"
                    If SNAP_status = "A" OR SNAP_status = "P" Then Full_case_list_array(7,m) = "P"
                    If HC_status = "A" OR HC_status = "P" Then Full_case_list_array(8,m) = "P"
                    If EA_status = "A" OR EA_status = "P" Then Full_case_list_array(9,m) = "P"
                    If GRH_status = "A" OR GRH_status = "P" Then Full_case_list_array(10,m) = "P"
                    If CCAP_status = "A" OR CCAP_status = "P" Then Full_case_list_array(12,m) = "P"

                    If cash_add_status = "A" OR cash_add_status = "P" Then Full_case_list_array(3,m) = cash_add_type
                    If cash_add_status = "A" OR cash_add_status = "P" Then Full_case_list_array(4,m) = "P"
                    If SNAP_add_status = "A" OR SNAP_add_status = "P" Then Full_case_list_array(7,m) = "P"
                    If HC_add_status = "A" OR HC_add_status = "P" Then Full_case_list_array(8,m) = "P"
                    If EA_add_status = "A" OR EA_add_status = "P" Then Full_case_list_array(9,m) = "P"
                    If GRH_add_status = "A" OR GRH_add_status = "P" Then Full_case_list_array(10,m) = "P"
                    If CCAP_add_status = "A" OR CCAP_add_status = "P" Then Full_case_list_array(12,m) = "P"

                    Full_case_list_array(11,m) = worker

					cash_add_status = ""		'BLANKING out these variables because they aren't read  on every loop and were carrying over.
					SNAP_add_status = ""
					HC_add_status = ""
					EA_add_status = ""
					GRH_add_status = ""
					CCAP_add_status = ""

                    'Redim Preserve Full_case_list_array (Ubound(Full_case_list_array,1), Ubound(Full_case_list_array,2)+1) 'Resize the array for the next case

                End If

                MAXIS_row = MAXIS_row + 1
                MAXIS_case_number = ""			'Blanking out variable
                m = m + 1
            Loop until MAXIS_row = 19
            PF8
            'Checking for the last page of cases.
            EMReadScreen last_page_check, 21, 24, 2	'because on REPT/PND2 it displays after trying to PF8, instead of when the second F8 is sent
        Loop until last_page_check = "THIS IS THE LAST PAGE"
    End if
next

k = 0	'Setting the inital value for the next array

For n = 0 to Ubound(Full_case_list_array,2)	'This will check all the cases from REPT/ACTV for any of the criteria selected in initial dialog
	MAXIS_case_number = Full_case_list_array(0,n)	'Setting the case number for the FuncLib fucntions to work
	IF SNAP_ABAWD_check = checked OR SNAP_UH_check = checked OR MFIP_tanf_check = checked OR child_only_mfip_check = checked OR mont_rept_check = checked OR HC_msp_check = checked OR adult_hc_check = checked OR family_hc_check = checked OR ltc_HC_check = checked OR waiver_HC_check = checked OR exclude_ievs_check = checked OR exclude_paris_check = checked then
		'////// Checking number of TANF months if requested
		IF MFIP_tanf_check = checked then
			IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" then
				STATS_counter = STATS_counter + 1
				navigate_to_MAXIS_screen "STAT", "TIME"
				EMReadScreen reg_mo, 2, 17, 69
				EMReadScreen ext_mo, 3, 19, 69
				If ext_mo = "   " then ext_mo = 0
				reg_mo = abs(reg_mo)
				ext_mo = abs(ext_mo)
				TANF_used = abs(reg_mo) + abs(ext_mo)
				PF3
			End If
		End If
		'////// Checking for adults on the MFIP grant if requested
		IF  child_only_mfip_check = checked then
			IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" then
				adult_on_mfip = False
				navigate_to_MAXIS_screen "ELIG", "MFIP"
                EMReadScreen look_for_version, 10, 24, 2
                If look_for_version <> "NO VERSION" Then
    				STATS_counter = STATS_counter + 1
    				EMReadScreen approval_check, 8, 3, 3
    				IF approval_check <> "APPROVED" then
    					EMReadScreen version_number, 1, 2, 12
    					prev_version = abs(version_number)-1
    					EMWriteScreen 0 & prev_version, 20, 79
    					transmit
    				End If
    				ReDim eligible_members_array (0)
    				ReDim non_mfip_members_array (0)
    				a = 0
    				b = 0
    				For row_to_check = 7 to 19
    					EMReadScreen pers_status, 10, row_to_check, 53
    					EMReadScreen memb_number, 2, row_to_check, 6
    					If pers_status = "INELIGIBLE" then
    						non_mfip_members_array(a) = memb_number
    						a = a + 1
    						ReDim Preserve non_mfip_members_array(a)
    					ElseIF pers_status = "ELIGIBLE  " then
    						eligible_members_array(b) = memb_number
    						b = b + 1
    						ReDim Preserve eligible_members_array(b)
    					Else
    						Exit For
    					End If
    				Next
    				navigate_to_MAXIS_screen "STAT", "MEMB"
    				For i = 0 to b
    					STATS_counter = STATS_counter + 1
    					EMWriteScreen eligible_members_array(i), 20, 76
    					transmit
    					EMReadScreen member_age, 2, 8, 76
    					If member_age = "  " then member_age = 0
    					If abs(member_age) > 18 then
    						adult_on_mfip = TRUE
    					ElseIF abs(member_age) = 18 AND eligible_members_array(i) = "01" THEN
    						adult_on_mfip = TRUE
    					End IF
    					If adult_on_mfip = TRUE then
    						Exit For
    					Else
    						adult_on_mfip = FALSE
    					End If
    				Next
                End If
			End If
		End If
		'//////Checking for monthly reporter
		IF mont_rept_check = checked then
			IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" then
				navigate_to_MAXIS_screen "ELIG", "MFIP"
                EMReadScreen look_for_version, 10, 24, 2
                If look_for_version <> "NO VERSION" Then
    				STATS_counter = STATS_counter + 1
    				EMReadScreen approval_check, 8, 3, 3
    				IF approval_check <> "APPROVED" then
    					EMReadScreen version_number, 1, 2, 12
    					prev_version = abs(version_number)-1
    					EMWriteScreen 0 & prev_version, 20, 79
    					transmit
    				End If
    				EMWriteScreen "MFSM", 20, 71
    				transmit
    				EMReadScreen reporter_type, 10, 8, 31
    				reporter_type = trim(reporter_type)
                End If
			End If
		End If
		'//////Checking for ABAWD Status
		IF SNAP_ABAWD_check = checked then
			IF Full_case_list_array(7,n) = "P" OR Full_case_list_array(7,n) = "A" then
				SNAP_with_ABAWD = False
				navigate_to_MAXIS_screen "ELIG", "FS"
				STATS_counter = STATS_counter + 1
				ReDim SNAP_HH_Array(0)
				c = 0
				For row_to_check = 7 to 19
					EMReadScreen pers_status, 10, row_to_check, 57
					EMReadScreen memb_number, 2, row_to_check, 10
					IF pers_status = "ELIGIBLE  " then
						SNAP_HH_Array(c) = memb_number
						c = c + 1
						ReDim Preserve SNAP_HH_Array(c)
					End If
				Next
				navigate_to_MAXIS_screen "STAT", "WREG"
				For j = 0 to c
					STATS_counter = STATS_counter + 1
					EMWriteScreen SNAP_HH_Array(j), 20, 76
					transmit
					EMReadScreen ABAWD_status, 2, 13, 50
					IF ABAWD_status = "10" OR ABAWD_status = "11" then
						SNAP_with_ABAWD = TRUE
						Exit For
					Else
						SNAP_with_ABAWD = FALSE
					End If
				Next
			End If
		End If
		'///// Determining if Case is Uncle Harry SNAP
		IF SNAP_UH_check = checked then
			IF Full_case_list_array(7,n) = "P" OR Full_case_list_array(7,n) = "A" then
				STATS_counter = STATS_counter + 1
				navigate_to_MAXIS_screen "ELIG", "FS"
				EMReadScreen type_of_SNAP, 13, 4, 3
				IF type_of_SNAP = "'UNCLE HARRY'" then
					UH_SNAP = TRUE
				Else
					UH_SNAP = FALSE
				End If
			End If
		End If
		'///// Finding if HC cases have Medicare Savings Programs active or pending
		IF HC_msp_check = checked then
			IF Full_case_list_array(8,n) = "A" OR Full_case_list_array(8,n) = "P" then
				STATS_counter = STATS_counter + 1
				navigate_to_MAXIS_screen "CASE", "CURR"
				'Determines if QMB is active
				Pending_MSP = False
				row = 1
				col = 1
				EMSearch "QMB:", row, col
				If row = 0 then
					QMB_active = FALSE
				Else
					EMReadScreen prog_status, 6, row, col + 5
					IF prog_status = "ACTIVE" OR prog_status = "APP OP" then
						QMB_active = TRUE
					ElseIf prog_status = "PENDIN" then
						Pending_MSP = TRUE
					End If
				End If
				'Determines if SLMB is active
				row = 1
				col = 1
				EMSearch "SLMB:", row, col
				If row = 0 then
					SLMB_active = FALSE
				Else
					EMReadScreen prog_status, 6, row, col + 6
					IF prog_status = "ACTIVE" OR prog_status = "APP OP" then
						SLMB_active = TRUE
					ElseIf prog_status = "PENDIN" then
						Pending_MSP = TRUE
					End If
				End If
				'Determines if QI1 is active
				row = 1
				col = 1
				EMSearch "Q1:", row, col
				If row = 0 then
					QI_active = FALSE
				Else
					EMReadScreen prog_status, 6, row, col + 4
					IF prog_status = "ACTIVE" OR prog_status = "APP OP" then
						QI_active = TRUE
					ElseIf prog_status = "PENDIN" then
						Pending_MSP = TRUE
					End If
				End If
				IF QMB_active = TRUE then
					MSP_actv = "QMB"
				ElseIf SLMB_active = TRUE then
					MSP_actv = "SLMB"
				ElseIF QI_active = TRUE then
					MSP_actv = "QI1"
				ElseIF Pending_MSP = TRUE then
					MSP_actv = "PEND"
				Else
					MSP_actv = "None"
				End If
			End If
		End If
		'////// Determining Family or Adult HC Cases
		If adult_hc_check = checked OR family_hc_check = checked then
			IF Full_case_list_array(8,n) = "A" or Full_case_list_array(8,n) = "P" then
				navigate_to_MAXIS_screen "ELIG", "HC"
				For row = 8 to 19
					EMReadScreen prog_status, 6, row, 50
					If prog_status = "ACTIVE" or prog_status = "PENDIN" then
						STATS_counter = STATS_counter + 1
						EMWriteScreen "x", 8, 26
						transmit
						EMReadScreen Elig_type, 2, 12, 72
						If Elig_type = "BT" OR Elig_type = "DT" then
							Specialty_HC = "TEFRA"
						ElseIF Elig_type = "09" OR Elig_type = "10" OR Elig_type = "25" then
							Specialty_HC = "Foster/IV-E"
						ElseIF Elig_type = "BC" then
							Specialty_HC = "SAGE/BC"
						ElseIf Elig_type = "11" OR Elig_type = "PX" OR Elig_type = "PC" OR Elig_type = "CB" OR Elig_type = "CK" OR Elig_type = "CX" OR Elig_type = "AA" then
							Family_HC = TRUE
						ElseIf Elig_type = "AX" OR Elig_type = "15" OR Elig_type = "16" OR Elig_type = "EX" OR Elig_type = "BX" OR Elig_type = "DX" OR Elig_type = "DP" OR Elig_type = "RM" then
							Adult_HC = TRUE
							Family_HC = FALSE
						End If
						If Specialty_HC <> "" OR Family_HC = TRUE then
							Adult_HC = FALSE
							Exit For
						End If
					End If
				Next
			End If
		End If
		'////// Determining LTC cases
		If ltc_HC_check = checked then
			IF Full_case_list_array(8,n) = "A" or Full_case_list_array(8,n) = "P" then
				navigate_to_MAXIS_screen "ELIG", "HC"
				STATS_counter = STATS_counter + 1
				EMWriteScreen "x", 8, 26
				transmit
				EMReadScreen hc_method, 1, 13, 76
				If hc_method = "L" then
					LTC_MA = TRUE
				Else
					LTC_MA = FALSE
				End If
			End If
		End If
		'////// Determining Waiver Cases
		If waiver_HC_check = checked then
			IF Full_case_list_array(8,n) = "A" or Full_case_list_array(8,n) = "P" then
				navigate_to_MAXIS_screen "ELIG", "HC"
				STATS_counter = STATS_counter + 1
				EMWriteScreen "x", 8, 26
				transmit
				EMReadScreen waiver_type, 1, 14, 76
				If waiver_type = "F" OR waiver_type = "G" OR waiver_type = "H" OR waiver_type = "I" OR waiver_type = "J" OR waiver_type = "K" OR waiver_type = "L" OR waiver_type = "M" OR waiver_type = "P" OR waiver_type = "Q" OR waiver_type = "R" OR waiver_type = "S" OR waiver_type = "Y" then
					Waiver_MA = TRUE
				Else
					Waiver_MA = FALSE
				End If
			End If
		End If
		'///// Determining if IEVS DAILs exist for this case
		IF exclude_ievs_check = checked then
			back_to_self
			EMWaitReady 0,0
			EMWriteScreen Full_case_list_array(0,n), 18, 43
			navigate_to_MAXIS_screen "DAIL", "DAIL"
			STATS_counter = STATS_counter + 1
			EMWriteScreen "x", 4, 12
			transmit
			EMWriteScreen " ", 7, 39
			EMWriteScreen "x", 12, 39
			transmit
			Do
				EMReadScreen msg_check, 11, 24, 2
				IF msg_check = "NO MESSAGES" THEN
					msg_check = ""
					Exit Do
				End IF
				ievs_dail_row = 1
				ievs_dail_col = 1
				EMSearch Full_case_list_array(0,n), ievs_dail_row, ievs_dail_col
				If ievs_dail_row = 0 then
					IEVS_DAIL = "N"
				Else
					IEVS_DAIL = "Y"
					Exit Do
				End If
				PF8
				EMReadScreen end_of_dail_check, 9, 24, 14
			Loop until end_of_dail_check = "LAST PAGE"
		End If
		'///// Determining if PARIS DAILs exist for this case
		IF exclude_paris_check = checked then
			back_to_self
			EMWaitReady 0,0
			EMWriteScreen Full_case_list_array(0,n), 18, 43
			navigate_to_MAXIS_screen "DAIL", "DAIL"
			STATS_counter = STATS_counter + 1
			EMWriteScreen "x", 4, 12
			transmit
			EMWriteScreen " ", 7, 39
			EMWriteScreen "x", 17, 39
			transmit
			Do
				EMReadScreen msg_check, 11, 24, 2
				IF msg_check = "NO MESSAGES" THEN
					msg_check = ""
					Exit Do
				End IF
				paris_dail_row = 1
				paris_dail_col = 1
				EMSearch Full_case_list_array(0,n), paris_dail_row, paris_dail_col
				If paris_dail_row = 0 then
					PARIS_DAIL = "N"
				Else
					PARIS_DAIL = "Y"
					Exit Do
				End If
				PF8
				EMReadScreen end_of_dail_check, 9, 24, 14
			Loop until end_of_dail_check = "LAST PAGE"
		End If
	End If

	'///// This is where the script determines which of the cases meet the criteria the user selected.
	'If Save_case_for_transfer is True once this Do Loop completes then the case information is saved for the transfer part in another array
	'This also determines which cases will be added to Excel
	Do	'The do loop is only here to be able to skip logic futher down in the list - it should never actually loop
		IF query_all_actv_check = checked Then
			Save_case_for_transfer = TRUE
			Exit Do 'IF the Query option is checked ALL cases get added to the list so none should have a FALSE
		End IF
        If all_cash_check = checked Then
            IF Full_case_list_array(3,n) = "CA" OR Full_case_list_array(5,n) = "CA" then Save_case_for_transfer = TRUE
        End If
		IF SNAP_check = checked then
			IF Full_case_list_array(7,n) = "P" OR Full_case_list_array(7,n) = "A" then Save_case_for_transfer = TRUE
		End If
		If mfip_check = checked then
			IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" then Save_case_for_transfer = TRUE
		End If
		If DWP_check = checked then
			IF Full_case_list_array(3,n) = "DW" OR Full_case_list_array(5,n) = "DW" then Save_case_for_transfer = TRUE
		End If

		If ga_check = checked then
			IF Full_case_list_array(3,n) = "GA" OR Full_case_list_array(5,n) = "GA" then Save_case_for_transfer = TRUE
		End If
		If msa_check = checked then
			IF Full_case_list_array(3,n) = "MS" OR Full_case_list_array(5,n) = "MS" then Save_case_for_transfer = TRUE
		End If
		If rca_check = checked then
			IF Full_case_list_array(3,n) = "RC" OR Full_case_list_array(5,n) = "RC" then Save_case_for_transfer = TRUE
		End If
		IF Full_case_list_array(9,n) = "P" AND EA_check = checked then Save_case_for_transfer = TRUE
		IF HC_check = checked then
			IF Full_case_list_array(8,n) = "A" OR Full_case_list_array(8,n) = "P" then Save_case_for_transfer = TRUE
		End If
		If GRH_check = checked then
			IF Full_case_list_array(10,n) = "A" OR Full_case_list_array(10,n) = "P" then Save_case_for_transfer = TRUE
		End If
		IF exclude_snap_check = checked then
			IF Full_case_list_array(7,n) = "A" OR Full_case_list_array(7,n) = "P" then Save_case_for_transfer = FALSE
		End if
		IF exclude_mfip_dwp_check = checked then
			IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" OR Full_case_list_array(3,n) = "DW" OR Full_case_list_array(5,n) = "DW" then Save_case_for_transfer = FALSE
		End if
		IF Full_case_list_array(9,n) = "P" AND exclude_ea_check = checked then Save_case_for_transfer = FALSE
		IF exclude_HC_check = checked then
			IF Full_case_list_array(8,n) = "A" OR Full_case_list_array(8,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF exclude_ga_msa_check = checked then
			IF Full_case_list_array(3,n) = "GA" OR Full_case_list_array(5,n) = "GA" OR Full_case_list_array(3,n) = "MS" OR Full_case_list_array(5,n) = "MS" then Save_case_for_transfer = FALSE
		End If
		IF exclude_grh_check = checked then
			IF Full_case_list_array(10,n) = "A" OR Full_case_list_array(10,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF exclude_RCA_check = checked then
			IF Full_case_list_array(3,n) = "RC" OR Full_case_list_array(5,n) = "RC" then Save_case_for_transfer = FALSE
		End If
		IF exclude_pending_check = checked then
			IF Full_case_list_array(7,n) = "P" OR Full_case_list_array(4,n) = "P" OR Full_case_list_array(6,n) = "P" OR Full_case_list_array(9,n) = "P" OR Full_case_list_array(8,n) = "P" OR Full_case_list_array(10,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF SNAP_Only_check = checked then
			IF Full_case_list_array(4,n) = "A" OR Full_case_list_array(4,n) = "P" OR Full_case_list_array(6,n) = "A" OR Full_case_list_array(6,n) = "P" OR Full_case_list_array(9,n) = "A" OR Full_case_list_array(9,n) = "P" OR Full_case_list_array(8,n) = "A" OR Full_case_list_array(8,n) = "P" OR Full_case_list_array(10,n) = "A" OR Full_case_list_array(10,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF HC_Only_check = checked then
			IF Full_case_list_array(4,n) = "A" OR Full_case_list_array(4,n) = "P" OR Full_case_list_array(6,n) = "A" OR Full_case_list_array(6,n) = "P" OR Full_case_list_array(9,n) = "A" OR Full_case_list_array(9,n) = "P" OR Full_case_list_array(7,n) = "A" OR Full_case_list_array(7,n) = "P" OR Full_case_list_array(10,n) = "A" OR Full_case_list_array(10,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF GRH_Only_check = checked then
			IF Full_case_list_array(4,n) = "A" OR Full_case_list_array(4,n) = "P" OR Full_case_list_array(6,n) = "A" OR Full_case_list_array(6,n) = "P" OR Full_case_list_array(9,n) = "A" OR Full_case_list_array(9,n) = "P" OR Full_case_list_array(8,n) = "A" OR Full_case_list_array(8,n) = "P" OR Full_case_list_array(7,n) = "A" OR Full_case_list_array(7,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF MFIP_Only_check = checked then
			IF Full_case_list_array(3,n) = "DW" OR Full_case_list_array(3,n) = "GA" OR Full_case_list_array(3,n) = "MS" OR Full_case_list_array(3,n) = "RC" OR Full_case_list_array(5,n) = "DW" OR Full_case_list_array(5,n) = "GA" OR Full_case_list_array(5,n) = "MS" OR Full_case_list_array(5,n) = "RC" OR Full_case_list_array(7,n) = "A" OR Full_case_list_array(7,n) = "P" OR Full_case_list_array(8,n) = "A" OR Full_case_list_array(8,n) = "P" OR Full_case_list_array(10,n) = "A" OR Full_case_list_array(10,n) = "P" OR Full_case_list_array(9,n) = "A" OR Full_case_list_array(9,n) = "P" then Save_case_for_transfer = FALSE
		End If
		IF SNAP_ABAWD_check = checked then
			IF SNAP_with_ABAWD = FALSE then Save_case_for_transfer = FALSE
		End If
		IF SNAP_UH_check = checked then
			IF UH_SNAP = FALSE then Save_case_for_transfer = FALSE
		End If
		IF MFIP_tanf_check = checked then
			IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" then
				IF TANF_used = "" then TANF_used = 0
				IF abs(TANF_used) < abs(tanf_months) then Save_case_for_transfer = FALSE
			End If
		End If
		IF child_only_mfip_check = checked AND adult_on_mfip = TRUE then Save_case_for_transfer = FALSE
		IF mont_rept_check = checked AND reporter_type <> "MONTHLY" then Save_case_for_transfer = FALSE
		IF HC_msp_check = checked AND MSP_actv = "None" then Save_case_for_transfer = FALSE
		IF adult_hc_check = checked AND Adult_HC = FALSE then Save_case_for_transfer = FALSE
		IF family_hc_check = checked AND Family_HC = FALSE then Save_case_for_transfer = FALSE
		IF ltc_HC_check = checked AND LTC_MA = FALSE then Save_case_for_transfer = FALSE
		IF waiver_HC_check = checked AND Waiver_MA = FALSE then Save_case_for_transfer = FALSE
		IF exclude_ievs_check = checked AND IEVS_DAIL = "Y" then Save_case_for_transfer = FALSE
		If exclude_paris_check = checked AND PARIS_DAIL = "Y" then Save_case_for_transfer = FALSE
		If Save_case_for_transfer <> TRUE THEN Save_case_for_transfer = FALSE
	Loop until Save_case_for_transfer <> ""

	'All_case_information_array is the big array with all of the information stored. These are the values of this array:
		'(0,#) - Case Number
		'(1,#) - Client Name
		'(2,#) = Review Date
		'(3,#) = Cash 1 Type
		'(4,#) = Cash 1 Status
		'(5,#) = Cash 2 Type
		'(6,#) = Cash 2 Status
		'(7,#) = TANF Used
		'(8,#) = Child Only MFIP status
		'(9,#) = SNAP Status
		'(10,#) = ABAWD on case?
		'(11,#) = Uncle Harry SNAP?
		'(12,#) = HC Status
		'(13,#) = Type of HC
		'(14,#) = Medicare Savings Prog
		'(15,#) = LTC MA?
		'(16,#) = Waiver MA?
		'(17,#) = Emergency Status
		'(18,#) = GRH Status
		'(19,#) = excel row to add information
		'(20,#) = IEVS DAIL?
		'(21,#) = PARIS DAIL?
		'(22,#) = Case transferred?
		'(23,#) = MFIP HRF?
		'(24,#) = CCAP Status

	IF Save_case_for_transfer = TRUE then
		'////// Add all information for qualifying cases into the Array
		All_case_information_array(0,k) = Full_case_list_array(0,n)
		All_case_information_array(1,k) = Full_case_list_array(1,n)
		All_case_information_array(2,k) = Full_case_list_array(2,n)
		All_case_information_array(3,k) = Full_case_list_array(3,n)
		All_case_information_array(4,k) = Full_case_list_array(4,n)
		All_case_information_array(5,k) = Full_case_list_array(5,n)
		All_case_information_array(6,k) = Full_case_list_array(6,n)
		All_case_information_array(7,k) = TANF_used
		IF Full_case_list_array(3,n) = "MF" OR Full_case_list_array(5,n) = "MF" then
			IF adult_on_mfip = FALSE then child_only = "Yes"
			IF adult_on_mfip = TRUE then child_only = "No"
		Else
			child_only = ""
		End If
		All_case_information_array(8,k) = child_only
		All_case_information_array(9,k) = Full_case_list_array(7,n)
		All_case_information_array(10,k) = SNAP_with_ABAWD
		All_case_information_array(11,k) = UH_SNAP
		All_case_information_array(12,k) = Full_case_list_array(8,n)
		IF Specialty_HC <> "" then
			All_case_information_array(13,k) = Specialty_HC
		ElseIf Family_HC = TRUE then
			All_case_information_array(13,k) = "Family"
		ElseIf Adult_HC = TRUE then
			All_case_information_array(13,k) = "Adult"
		End IF
		All_case_information_array(14,k) = MSP_actv
		All_case_information_array(15,k) = LTC_MA
		All_case_information_array(16,k) = Waiver_MA
		All_case_information_array(17,k) = Full_case_list_array(9,n)
		All_case_information_array(18,k) = Full_case_list_array(10,n)
		All_case_information_array(19,k) = excel_row
		All_case_information_array(20,k) = IEVS_DAIL
		All_case_information_array(21,k) = PARIS_DAIL
		IF reporter_type = "MONTHLY" THEN
			All_case_information_array(23,k) = "Y"
		ElseIf reporter_type = "" then
			All_case_information_array(23,k) = ""
		Else
			All_case_information_array(23,k) = "N"
		End IF
		All_case_information_array(24,k) = Full_case_list_array(12,n)

		'///// Resizing the storage array for the next loop
		Redim Preserve All_case_information_array (UBound(All_case_information_array,1), UBound(All_case_information_array,2)+1)

		'ADD THE INFORMATION TO XCEL HERE
		ObjExcel.Cells(excel_row, 1).Value = Full_case_list_array(11,n)
		ObjExcel.Cells(excel_row, 2).Value = All_case_information_array(0,k)
		ObjExcel.Cells(excel_row, 3).Value = All_case_information_array(1,k)
		ObjExcel.Cells(excel_row, 4).Value = All_case_information_array(2,k)
		'ObjExcel.Cells(excel_row, 5).Value = abs(days_pending)
		ObjExcel.Cells(excel_row, snap_actv_col).Value = All_case_information_array(9,k)
		IF SNAP_ABAWD_check = checked THEN ObjExcel.Cells(excel_row, ABAWD_actv_col). Value = All_case_information_array(10,k)
		IF SNAP_UH_check = checked THEN ObjExcel.Cells(excel_row, UH_actv_col).Value = All_case_information_array(11,k)
		ObjExcel.Cells(excel_row, cash_one_prog_col).Value = All_case_information_array(3,k)
		ObjExcel.Cells(excel_row, cash_one_actv_col).Value = All_case_information_array(4,k)
		ObjExcel.Cells(excel_row, cash_two_prog_col).Value = All_case_information_array(5,k)
		ObjExcel.Cells(excel_row, cash_two_actv_col).Value = All_case_information_array(6,k)
		IF MFIP_tanf_check = checked THEN ObjExcel.Cells(excel_row, TANF_mo_col).Value = All_case_information_array(7,k)
		IF child_only_mfip_check = checked THEN ObjExcel.Cells(excel_row, child_only_col).Value = All_case_information_array(8,k)
		IF mont_rept_check = checked THEN ObjExcel.Cells(excel_row, mont_rept_col).Value = All_case_information_array(23,k)
		IF ccap_check = checked OR exclude_ccap_check = checked THEN ObjExcel.Cells(excel_row,ccap_col) = All_case_information_array(24,k)
		ObjExcel.Cells(excel_row, hc_actv_col).Value = All_case_information_array(12,k)
		IF adult_hc_check = checked OR family_hc_check =checked THEN ObjExcel.Cells(excel_row, hc_type_col).Value = All_case_information_array(13,k)
		IF HC_msp_check = checked THEN ObjExcel.Cells(excel_row, MSP_actv_col).Value = All_case_information_array(14,k)
		IF ltc_HC_check = checked THEN ObjExcel.Cells(excel_row, LTC_col).Value = All_case_information_array(15,k)
		IF waiver_HC_check = checked THEN ObjExcel.Cells(excel_row, Waiver_col).Value = All_case_information_array(16,k)
		ObjExcel.Cells(excel_row, EA_actv_col).Value = All_case_information_array(17,k)
		ObjExcel.Cells(excel_row, GRH_actv_col).Value = All_case_information_array(18,k)
		IF exclude_ievs_check = checked THEN ObjExcel.Cells(excel_row, ievs_col).Value = All_case_information_array(20,k)
		IF exclude_paris_check = checked THEN ObjExcel.Cells(excel_row, paris_col).Value = All_case_information_array(21,k)
		excel_row = excel_row + 1
		k = k + 1 'Goes to the next entry for the All_case_information_array
	End if
	'Blanking out variables for next go round
	MAXIS_case_number = ""
	Save_case_for_transfer = ""
	reg_mo = ""
	ext_mo = ""
	TANF_used = ""
	ReDim eligible_members_array (0)
	ReDim non_mfip_members_array (0)
	adult_on_mfip = ""
	reporter_type = ""
	SNAP_with_ABAWD = ""
	ReDim SNAP_HH_Array(0)
	UH_SNAP = ""
	Pending_MSP = ""
	QMB_active = ""
	SLMB_active = ""
	QI_active = ""
	MSP_actv = ""
	Specialy_HC = ""
	Family_HC = ""
	Adult_HC = ""
	LTC_MA = ""
	Waiver_MA = ""
	IEVS_DAIL = ""
	PARIS_DAIL = ""
	If case_found_limit <> "" Then
		If k = case_found_limit Then Exit For
	End If
Next

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time


'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Totals up the number of cases found so the user has a count
cases_found = abs(UBound(All_case_information_array,2))
cases_to_xfer_numb = cases_found

'Stops the script before transfer when the query option is selected.
IF query_all_actv_check = checked THEN
	'Adding the number of cases count and formatting the speadsheet
	objExcel.Cells(4, col_to_use - 1).Font.Bold = TRUE
	ObjExcel.Cells(4, col_to_use - 1).Value = "Number of cases that meet selected criteria:"	'Goes back one, as this is on the next row
	ObjExcel.Cells(4, col_to_use).Value = cases_found

	'Autofitting columns
	For col_to_autofit = 1 to col_to_use
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next
	'Logging usage stats
	script_end_procedure("Success! The script is complete. " & vbCr & cases_found & " cases have been found in your selected case loads." & vbCr & "Your Excel Sheet has all the information about these cases.")
End If

'Second DIalog needs to be after the calculations so the variables have value
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 376, 130, "Select Transfer Options"
  CheckBox 5, 30, 130, 10, "Check here to have the script transfer", transfer_check
  EditBox 140, 25, 20, 15, cases_to_xfer_numb
  EditBox 230, 25, 135, 15, worker_receiving_cases
  CheckBox 5, 45, 185, 10, "Check here to have a case note entered for each case", case_note_check
  EditBox 85, 60, 80, 15, worker_signature
  EditBox 85, 80, 80, 15, new_worker
  CheckBox 5, 100, 185, 10, "Check here to have a MEMO sent for each case", memo_check
  CheckBox 5, 115, 175, 10, "Check here if you do not want to transfer any cases", query_check
  ButtonGroup ButtonPressed
    OkButton 265, 110, 50, 15
    CancelButton 320, 110, 50, 15
  Text 60, 10, 55, 10, "The script found"
  Text 115, 10, 20, 10, cases_found
  Text 140, 10, 130, 10, "cases that meet your selected criteria"
  Text 165, 30, 60, 10, "of these cases to:"
  Text 230, 40, 140, 25, "Enter the entire 7-digit number x1 number. You may enter more than one worker, seperate workers by a comma."
  Text 15, 85, 70, 10, "New Worker's Name"
  Text 15, 65, 65, 10, "Sign your case note"
  Text 170, 80, 170, 20, "This is optional, it only adds the worker's name to the case note - you can only enter one name."
EndDialog

'Running the dialog to get transfer information
Do
	Do
		Dialog Dialog1
		cancel_confirmation
		err_msg = ""
		If cases_to_xfer_numb = "" THEN cases_to_xfer_numb = 0
		IF transfer_check = checked AND worker_receiving_cases = "" then err_msg = err_msg & vbCR & "You must enter a worker number to transfer cases to"
		IF abs(cases_to_xfer_numb) > abs(cases_found) then err_msg = err_msg & vbCr & "You cannot transfer more cases than were found to transfer"
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		If err_msg <> "" then MsgBox err_msg
	Loop until err_msg = ""
	IF transfer_check = unchecked AND query_check = unchecked THEN MsgBox "You must select an option"
	IF transfer_check = checked AND query_check = checked THEN MsgBox "You cannot select both"
Loop until transfer_check = checked OR query_check = checked

cases_to_xfer_numb = abs(cases_to_xfer_numb)	'Sometimes the script thinks this is a string and does not do math correctly.

'Standardizing the worker numbers to uppercase so that transfering doesn't break
worker_receiving_cases = UCase(worker_receiving_cases)
worker_receiving_cases = Replace(worker_receiving_cases, " ", "")	'Removing stray spaces
'Creating the array of all workers to receive cases
receiving_worker_array = split(worker_receiving_cases, ",")

r = 0 	'counter for the receiving worker array
P = 0 	'Counter for the cases transferred
'Transfering the cases
If transfer_check = checked then
	Do
		back_to_self
		MAXIS_case_number = All_case_information_array(0,p) 'setting case number variable for the FuncLib functions to work
		navigate_to_MAXIS_screen "SPEC", "XFER"
		STATS_counter = STATS_counter + 1
		EMWriteScreen "x", 7, 16
		transmit
		PF9
		EMWriteScreen receiving_worker_array(r), 18, 61
		transmit
		EMReadScreen confirm_xfer, 4, 24, 2
		IF confirm_xfer <> "CASE" then
			'If a transfer is not successful it will be noted on the spreadsheet and a msgbox will alert the user but the script will not stop.
			'Option to disable the message box if this holds up the runtime
			MsgBox "This case " & MAXIS_case_number & " cannot be transferred and has been noted on the spreadsheet"
			PF10
			ObjExcel.Cells (All_case_information_array(19,p), xfered_col).Value = "N"
		ElseIf confirm_xfer = "CASE" then
			ObjExcel.Cells (All_case_information_array(19,p), xfered_col).Value = "Y"
			ObjExcel.Cells (All_case_information_array(19,p), new_worker_col).Value = receiving_worker_array(r)
			total_cases_transfered = total_cases_transfered + 1 	'This counts the successful transfers
			r = r + 1 'The cases are assigned to multiple workers on a basic rotation
			IF r > UBound(receiving_worker_array) THEN r = 0
			IF case_note_check = checked then
				'Writes a case note if requested.
				Call start_a_blank_case_note
				Call write_variable_in_case_note ("***Case Transfer within County***")
				Call write_bullet_and_variable_in_case_note ("Case Transferred to", new_worker)
				Call write_variable_in_case_note ("Transfered by bulk script")
				IF memo_check = checked then Call write_variable_in_case_note ("Memo sent to clt of transfer")
				Call write_variable_in_case_note ("---")
				Call write_variable_in_case_note (worker_signature)
				case_note_check = checked 'adding this because sometimes the loop loses this value for some reason
			End If
			IF memo_check = checked then
				Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)

				If memo_opened = True Then
					Call write_variable_in_SPEC_MEMO ("*** This is just an informational notice ***")
					Call write_variable_in_SPEC_MEMO ("Your case has been transferred.")
					Call write_variable_in_SPEC_MEMO ("I will be your new case worker.")
					Call write_variable_in_SPEC_MEMO ("   ")
					Call write_variable_in_SPEC_MEMO ("This is not a request for any information.")
					Call write_variable_in_SPEC_MEMO ("If I need anything from you, I will send a separate request")
					Call write_variable_in_SPEC_MEMO ("   ")
					Call write_variable_in_SPEC_MEMO ("Thank you")
					PF4
					PF3
				End If
				memo_check = checked 'adding this because sometimes the loop loses this value for some reason

			End If
			'If cases_to_xfer_numb = total_cases_transfered Then Exit Do
		End If
		MAXIS_case_number = "" 'Blanking out variable
		p = p + 1
		If p = UBound(All_case_information_array,2) Then Exit Do
	Loop until total_cases_transfered = cases_to_xfer_numb 'continues to attempt to transfer until the requested number is reached
End If

If total_cases_transfered = "" then total_cases_transfered = 0

'Adding some counts to the excel sheet
objExcel.Cells(4, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(5, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(4, col_to_use - 1).Value = "Number of cases that meet selected criteria:"	'Goes back one, as this is on the next row
ObjExcel.Cells(4, col_to_use).Value = cases_found
ObjExcel.Cells(5, col_to_use - 1).Value = "Number of cases transferred"	'Goes back one, as this is on the next row
ObjExcel.Cells(5, col_to_use).Value = total_cases_transfered

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1
'Logging usage stats
script_end_procedure("Success! The script is complete. " & vbCr & total_cases_transfered & " cases have been transferred." & vbCr & "Your Excel Sheet has all the detail")
