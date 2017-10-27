'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - COUNTED ABAWD MONTHS.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 600         'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("10/20/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog ABAWD_dialog, 0, 0, 176, 125, "Counted ABAWD months"
  EditBox 85, 65, 50, 15, MAXIS_case_number
  EditBox 85, 85, 20, 15, HH_memb
  ButtonGroup ButtonPressed
    OkButton 40, 105, 50, 15
    CancelButton 95, 105, 50, 15
  Text 35, 70, 50, 10, "Case Number:"
  Text 15, 20, 150, 35, "This script will provide information regarding public assistance issuanceson the case, and what is marked on the ABAWD tracking record for each member."
  GroupBox 10, 5, 160, 55, "Using this script:"
  Text 45, 90, 35, 10, "Member #:"
EndDialog

'The script============================================================================================================================
'Connects to MAXIS, grabbing the case MAXIS_case_number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
HH_memb = "01"

'establishing what MAXIS_footer_month and year are for WREG panel/ATR months determination
MAXIS_footer_month 	= CM_mo
MAXIS_footer_year 	= CM_yr

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog ABAWD_dialog				'main dialog
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "* Enter a valid case number."		'mandatory field
		IF len(HH_memb) <> 2 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "* Enter a valid 2-digit member number."		'mandatory field
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Creating a custom dialog for determining who the HH members are
'call HH_member_custom_dialog(HH_member_array)

'Will check ATR in current month/year
back_to_self
EMWriteScreen CM_mo, 20, 43				'
EMWriteScreen CM_yr, 20, 46

'For each HH_memb in HH_member_array
Call navigate_to_MAXIS_screen("STAT", "WREG")

'Checking for PRIV cases.
EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
IF priv_check = "PRIVIL" THEN script_end_procedure("This case is a privliged case. You do not have access to this case.")

Call write_value_and_transmit(HH_memb, 20, 76)

EMReadScreen wreg_total, 1, 2, 78
If wreg_total = "0" then script_end_procedure("WREG panel does not exist for this member. The script will now end.")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'adding column header information to the Excel list
ObjExcel.Cells(1, 1).Value = "Month"
ObjExcel.Cells(1, 2).Value = "MEMB 01"
ObjExcel.Cells(1, 3).Value = "SNAP"
ObjExcel.Cells(1, 4).Value = "GA"
ObjExcel.Cells(1, 5).Value = "MFIP"
ObjExcel.Cells(1, 6).Value = "MF - FS"
ObjExcel.Cells(1, 7).Value = "DWP"
ObjExcel.Cells(1, 8).Value = "RCA"
ObjExcel.Cells(1, 9).Value = "MSA"

'formatting the cells
'FOR i = 1 to col_to_use
FOR i = 1 to 9
	objExcel.Cells(1, i).Font.Bold = True		'bold font
	objExcel.Columns(i).AutoFit()				'sizing the columns
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
NEXT
	
excel_row = 2

EmWriteScreen "x", 13, 57		'Pulls up the WREG tracker'
transmit
EMREADScreen tracking_record_check, 15, 4, 40  		'adds cases to the rejection list if the ABAWD tracking record cannot be accessed.
If tracking_record_check <> "Tracking Record" then script_end_procedure("Unable to enter ABAWD tracking record of member.")
bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
bene_yr_row = 10
DO
    'establishing variables for specific ABAWD counted month dates
    If bene_mo_col = "19" then counted_date_month = "01"
    If bene_mo_col = "23" then counted_date_month = "02"
    If bene_mo_col = "27" then counted_date_month = "03"
    If bene_mo_col = "31" then counted_date_month = "04"
    If bene_mo_col = "35" then counted_date_month = "05"
    If bene_mo_col = "39" then counted_date_month = "06"
    If bene_mo_col = "43" then counted_date_month = "07"
    If bene_mo_col = "47" then counted_date_month = "08"
    If bene_mo_col = "51" then counted_date_month = "09"
    If bene_mo_col = "55" then counted_date_month = "10"
    If bene_mo_col = "59" then counted_date_month = "11"
    If bene_mo_col = "63" then counted_date_month = "12"
	
    'counted date year: this is found on rows 7-11. Row 11 is current year plus one, so this will be exclude this list.
    If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", date), 2)
    If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, date)), 2)
    If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, date)), 2)
    If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)
    	
	EMReadScreen counted_date_year, 2, bene_yr_row, 14								'reading counted year date
	abawd_counted_months_string = counted_date_month & "/" & counted_date_year		'creating new date variable
	'abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
	
	ObjExcel.Cells(excel_row, 1).Value = abawd_counted_months_string
	
	'reading to see if a month is counted month or not
	EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
	IF is_counted_month <> "_" then ObjExcel.Cells(excel_row, 2).Value = is_counted_month
	excel_row = excel_row + 1
	
	bene_mo_col = bene_mo_col - 4		're-establishing serach once the end of the row is reached
	IF bene_mo_col = 15 THEN
		bene_yr_row = bene_yr_row - 1
		bene_mo_col = 63
	END IF
LOOP until bene_yr_row = 6	
	
PF3 	'to exit the ABAWD tracking record
	
'----------------------------------------------------------------------------------------------------INQX
INQX_yr = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)

Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWritescreen "01", 6, 38
EMWritescreen INQX_yr, 6, 41
EMWritescreen CM_mo, 6, 53
EMwritescreen CM_yr, 6, 56
EMWritescreen "X", 9, 5		'Snap
EMWritescreen "X", 10, 5	'MFIP
EMWritescreen "X", 11, 5 	'GA	
EMWritescreen "X", 15, 5	'RCA
EMWritescreen "X", 13, 50	'MSA
EMWritescreen "X", 17, 50 	'DWP
transmit

'checking to see if HG has been issued for the month selected: MONY/INQX----------------------------------------------------------------------------------------------------

EMReadScreen no_issuance, 11, 24, 2 
If no_issuance = "NO ISSUANCE" then script_end_procedure(HH_memb & " does not have any issuance during this period. The script will now end.")
EMReadScreen single_page, 8, 17, 73
If trim(single_page) = "" then one_page = True

'Checks for cases with more then 9 pages of issuances
Do 
	PF8
	EMReadScreen page_check, 21, 24, 2 
	If page_check = "CAN NOT PAGE THROUGH " then script_end_procedure("Case has more than 9 pages of issuance. Please process/review manually.")
Loop until page_check = "THIS IS THE LAST PAGE"

'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
Do
	PF7
	EMReadScreen first_page_check, 20, 24, 2
LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page

Excel_row = 2
DO
	row = 6				'establishing the row to start searching for issuance
	tracking_month = objExcel.cells(excel_row, 1).Value	're-establishing the case number to use for the case 
	If trim(tracking_month) = "" then exit do 
	
	Do 
	    Do    
	    	EMReadScreen issuance_month, 2, row, 73
	    	EMReadScreen issuance_year, 2, row, 79
			EMReadScreen issuance_day, 2, row, 65
	    	INQX_issuance = issuance_month & "/" & issuance_year
	    	If trim(INQX_issuance) = "" then exit do 
	    	
	    	If tracking_month = INQX_issuance then 	
	    		EMReadScreen prog_type, 5, row, 16		
	    		prog_type = trim(prog_type)
	    		EMReadScreen amt_issued, 7, row, 40
				If issuance_day <> "01" then amt_issued = amt_issued & "*"
	    		If prog_type = "FS" 	then fs_issued = fs_issued + amt_issued
	    		If prog_type = "GA" 	then ga_issued = ga_issued + amt_issued
	    		If prog_type = "MF-MF" 	then mfip_issued = mfip_issued + amt_issued
	    		If prog_type = "MF-FS" 	then mffs_issued = mffs_issued + amt_issued
	    		If prog_type = "DW" 	then dw_issued = dw_issued + amt_issued
	    		If prog_type = "RC" 	then rc_issued = rc_issued + amt_issued
	    		If prog_type = "MS" 	then ms_issued = ms_issued + amt_issued
	    	End if 
	    	row = row + 1
	    Loop until row = 18
		If one_page = True then exit do 
		PF8
		EMReadScreen last_page_check, 21, 24, 2
		If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"
	
	ObjExcel.Cells(excel_row, 3).Value = fs_issued
	ObjExcel.Cells(excel_row, 4).Value = ga_issued
	ObjExcel.Cells(excel_row, 5).Value = mfip_issued
	ObjExcel.Cells(excel_row, 6).Value = mffs_issued
	ObjExcel.Cells(excel_row, 7).Value = dw_issued
	ObjExcel.Cells(excel_row, 8).Value = rc_issued
	ObjExcel.Cells(excel_row, 9).Value = ms_issued
		
	amt_issued = ""
	fs_issued = ""
	ga_issued = ""
	mfip_issued = ""
	mffs_issued = ""
	dw_issued = ""
	rc_issued = ""
	ms_issued = ""
	
	If one_page <> True then 
	    'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
	    Do
	    	PF7
	    	EMReadScreen first_page_check, 20, 24, 2
	    LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page
	End if 
	
	excel_row = excel_row + 1
Loop

FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

script_end_procedure("Success, please review the ABAWD's information.")
	