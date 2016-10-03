'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - COPY MAXIS CASE DATA TO EXCEL.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 38                      'manual run time in seconds  this run time only includes opening the spreadsheet, copying the template, and renaming it...more to come...
STATS_denomination = "C"       'I is for each case
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

'CONSTANTS=============================
'These are the start rows for each MAXIS panel
memb_row = 6
memi_row = 19
addr_row = 23
revw_row = 42
abps_row = 44
acct_row = 46
acut_row = 62
bils_row = 80
busi_row = 143
cars_row = 177
cash_row = 192
coex_row = 193
dcex_row = 210
dfln_row = 237
diet_row = 249
disa_row = 261	
dstt_row = 277
eats_row = 280
emma_row = 285
emps_row = 290
faci_row = 301
fmed_row = 308
hest_row = 337
imig_row = 345
insa_row = 361
jobs1_row = 370
jobs2_row = 377
jobs3_row = 384
medi_row = 391
mmsa_row = 398
msur_row = 402
othr_row = 403
pare_row = 415
pben1_row = 433
pben2_row = 439
pben3_row = 445
pded_row = 451
preg_row = 465
rbic_row = 470
rest_row = 495
schl_row = 506
secu_row = 516
shel_row = 531
sibl_row = 566
spon_row = 569
stec_row = 573
stin_row = 585
stwk_row = 595
unea1_row = 607
unea2_row = 613
unea3_row = 619
wkex_row = 625
wreg_row = 674

'custom function for determining data validation values in the given cell and matching with what is needed...
FUNCTION check_for_data_validation(cell_row, cell_column, maxis_value, objExcel, objWorkbook, objTemplate, objNewSheet)
	'backing out of the function and skipping that cell if the script finds a "?" and notifying the user of the problem...
	IF InStr(maxis_value, "?") <> 0 THEN 
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "The script is attempting to write a ''?'' to the template. This value is not supported. The script will skip this value for the cell at row " & cell_row & " and column " & cell_column & ".", vbInformation + vbSystemModal, "Invalid Character Found"
		EXIT FUNCTION
	END IF
	
	'creating string of alphabet for comparison
	alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	'converting the row and column
	specific_cell = Mid(alphabet, cell_column, 1) & cell_row

	'creating new object for that specific cell
	SET objTempRange = objNewSheet.Range(specific_cell)
	'grabbing the controls range for the data validation on that cell
	source_range = objTempRange.Validation.Formula1
	
	'Checking to see that the data validation is from a reference to the controls sheet or if it has been hard-coded to the field
	IF InStr(source_range, "controls") = 0 THEN 					'<<< if the values are hard coded
		all_possible_values = split(source_range, ",")				'<<< looking for the specific value in the list of values
		FOR EACH possible_value IN all_possible_values
			IF InStr(possible_value, maxis_value) = 1 THEN 			'<<< if the value is found, use that entire value for the cell_row
				objExcel.Cells(cell_row, cell_column).Value = possible_value
				EXIT FOR											'<<< and EXIT the FOR/NEXT
			END IF
		NEXT
	
	ELSE			'<<< if the values are taken from 'controls'
		'Trimming and modifiying the source range to make it workable
		source_range = replace(source_range, "$", "")
		source_range = replace(source_range, "=controls!", "")
		
		'Determining the start and end of the range
		colon_pos = InStr(source_range, ":")
		start_cell = left(source_range, colon_pos - 1)
		end_cell = right(source_range, len(source_range) - len(start_cell) - 1)
	
		'Converting the range to script-friendly parts...
		'...the start row and column
		FOR i = 1 TO len(start_cell)
			IF IsNumeric(right(start_cell, i)) = FALSE THEN 
				start_row = right(start_cell, i - 1)
				start_col = left(start_cell, len(start_cell) - len(start_row))
				EXIT FOR
			END IF
		NEXT
		
		'...and the end row
		FOR i = 1 TO len(end_cell)
			IF IsNumeric(right(end_cell, i)) = FALSE THEN 
				end_row = right(end_cell, i - 1)
				end_col = left(end_cell, len(end_cell) - len(end_row))
				EXIT FOR
			END IF
		NEXT
			
		'Grabbing validation values from control worksheet
		SET objControls = objWorkbook.Worksheets("Controls")
		
		'now going through the range of validation cells to find the one that matches what we found in MAXIS
		FOR i = (start_row * 1) TO (end_row * 1)
			control_col = InStr(alphabet, start_col)
			IF InStr(objControls.Cells(i, control_col).Value, maxis_value) = 1 THEN 
				objExcel.Cells(cell_row, cell_column).Value = objControls.Cells(i, control_col).Value
				EXIT FOR
			END IF
		NEXT
	END IF
END FUNCTION
'============================================================================================

'DIALOGS------------------------------------------------------------------------------------------------------------------
'This is the dialog that allows the user to open the existing spreadsheet
DO
    DO
        BeginDialog Dialog1, 0, 0, 381, 320, "MAXIS Case Replicator 9000, Version 1.0"
          Text 10, 10, 200, 10, "Hello, human. Welcome to the MAXIS Case Replicator 9000."
          Text 10, 25, 365, 20, "This script works in conjunction with the training case generator, to grab case information out of MAXIS and insert it into a new scenario in scenario spreadsheet."
          Text 10, 50, 365, 25, "Because this script relies on the training case generator template, it has a few limitations that you should be aware of before using it. Please take a moment to review the limitations below to ensure you are getting the most out of this script."
          Text 10, 275, 140, 10, "Select an Excel file for training scenarios:"
          Text 30, 185, 345, 10, "* Regarding JOBS -- this script will only be able to pull up to 3 active JOBS panels per client."
          Text 30, 80, 345, 25, "* To protext client privacy and data, if you are grabbing data from PRODUCTION or INQUIRY, you will be required to rename individuals in the scenario. If you are grabbing data from TRAINING, you will have the option of renaming the individuals in the scenario."
          Text 30, 125, 345, 10, "* Regarding BUSI -- this script will only grab data from the first BUSI panel per client."
          Text 30, 110, 345, 10, "* Regarding ACCT -- this script will only grab data from the first ACCT panel per client."
          Text 30, 140, 345, 10, "* Regarding CARS -- this script will only grab data from the first CARS panel per client."
          Text 30, 155, 345, 25, "* Regarding FACI -- this script will only grab data from the first FACI panel per client. Additionally, it will only grab the first begin and end data for that facility. If the client has multiple entries and exits from that facility, it will not be captured by this script."
          Text 30, 200, 345, 10, "* Regarding UNEA -- this script will only be able to pull up to 3 active UNEA panels per client."
          Text 30, 215, 345, 25, "* Miscellaneous -- there may be additional data points in the case you are copying that are not picked up by this script. This is a limitation of the training case generator as it currently exists. This script will only pick up on data points that the training case generator is able to write to MAXIS."
          EditBox 150, 270, 175, 15, training_case_creator_excel_file_path
          ButtonGroup ButtonPressed
            PushButton 330, 270, 45, 15, "Browse...", select_a_file_button
            OkButton 270, 300, 50, 15
            CancelButton 325, 300, 50, 15
        EndDialog
    
    	Dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(training_case_creator_excel_file_path, ".xlsx")
    Loop until ButtonPressed = OK and training_case_creator_excel_file_path <> ""
    
    'checking that MX is not timed out
    CALL check_for_MAXIS(false)
    
    'opening the spreadsheet
    SET objExcel = CreateObject("Excel.Application")
    objExcel.Visible = TRUE
    SET objWorkbook = objExcel.Workbooks.Add(training_case_creator_excel_file_path)
    objExcel.DisplayAlerts = FALSE
	
	'Asking the user to confirm the spreadsheet
	confirm_spreadsheet = MsgBox ("Is this the correct spreadsheet? Press YES to confirm and continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel, vbQuestion + vbSystemModal, "Confirm SpreadSheet")
	IF confirm_spreadsheet = vbCancel THEN script_end_procedure("Script cancelled.")
	IF confirm_spreadsheet = vbNo THEN 
		objWorkbook.Close
		objExcel.Quit
	END IF
LOOP UNTIL confirm_spreadsheet = vbYes

'naming the scenario	 
DO
	err_msg = ""
	
	BeginDialog Dialog1, 0, 0, 216, 60, "Scenario Creator"
	EditBox 105, 10, 105, 15, scenario_name
	ButtonGroup ButtonPressed
		OkButton 110, 40, 50, 15
		CancelButton 160, 40, 50, 15
	Text 10, 15, 90, 10, "Name your new scenario:"
	EndDialog
	
	DIALOG
	cancel_confirmation
	IF scenario_name = "" THEN err_msg = err_msg & vbCr & "* You must give the scenario a name."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'double checking that MX is not timed out
CALL check_for_MAXIS(false)

objExcel.WorkSheets("Template").Activate				
SET objTemplate = objWorkbook.Worksheets("Template")			'activating the template worksheet
SET objTemplateRange = objTemplate.Range("A1:Z700")				'selecting everything relevant to Krabappel
objTemplateRange.Copy											'copying the range

SET objNewSheet = objWorkbook.Sheets.Add()			'giving the new worksheet a new name
objNewSheet.Name = scenario_name
objNewSheet.Paste									'pasting in the template

'autofitting the columns
FOR column = 1 to 26
	objExcel.Columns(column).AutoFit()
NEXT

'saving the workbook
objWorkbook.SaveAs(training_case_creator_excel_file_path) 

'The script ------- grabbing the MAXIS case number
EMConnect ""

'grabbing the MAXIS case number already active
CALL MAXIS_case_number_finder(MAXIS_case_number)

DO
    BeginDialog Dialog1, 0, 0, 211, 60, "Confirm MAXIS Case Number"
      EditBox 140, 10, 65, 15, MAXIS_case_number
      ButtonGroup ButtonPressed
        OkButton 105, 40, 50, 15
        CancelButton 155, 40, 50, 15
      Text 10, 15, 125, 10, "Please enter a case number to copy:"
    EndDialog
    
    DIALOG
    cancel_confirmation
	IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "Please enter a valid MAXIS case number."
LOOP UNTIL 

'...and now we burgle...
back_to_SELF
CALL find_variable("Environment: ", maxis_enviro, 5)

'navigating to MEMB...
'the script is going to grab the REF number, the date of birth, the age, gender, etc from MEMB
'it will also navigate to the next member, building an array of ref numbers as it goes
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
excel_col = 3
client_array = ""
DO
	'Reading the bits off MEMB
	STATS_manualtime = STATS_manualtime + 45 '<<< 44.96 seconds to read and write all data points to the Excel file per client
	EMReadScreen ref_num, 2, 4, 33
	EMReadScreen client_dob, 5, 8, 42
		client_dob = replace(client_dob, " ", "/")
	EMReadScreen client_dob_verif, 2, 8, 68
	EMReadScreen client_age, 2, 8, 76
	EMReadScreen client_gender, 1, 9, 42
	EMReadScreen id_verif, 2, 9, 68
	EMReadScreen relation_to_app, 2, 10, 42
		relation_to_app = trim(relation_to_app)
		relation_to_app = replace(relation_to_app, "  ", " - ")
	EMReadScreen client_language, 20, 12, 42
		client_language = replace(client_language, "_", "")
		client_language = replace(client_language, "  ", " - ")
	EMReadScreen needs_interp, 1, 14, 68
	'EMReadScreen cl_alias, 1, 15, 42				' <<< currently (7/11/2016) commented out... Krabappel is not capable of handling aliases...going to hard code to "N"
	cl_alias = "N"
	EMReadScreen hisp_latino, 1, 16, 68
	
	objExcel.Cells(2, excel_col).Value = ref_num
	'writing the bits from MEMB
	objExcel.Cells(memb_row + 3, excel_col).Value = client_dob
	objExcel.Cells(memb_row + 4, excel_col).Value = client_age
	objExcel.Cells(memb_row + 5, excel_col).Value = client_dob_verif
	objExcel.Cells(memb_row + 6, excel_col).Value = client_gender
	objExcel.Cells(memb_row + 7, excel_col).Value = id_verif
	CALL check_for_data_validation(memb_row + 8, excel_col, relation_to_app, objExcel, objWorkbook, objTemplate, objNewSheet)
	objExcel.Cells(memb_row + 9, excel_col).Value = client_language
	objExcel.Cells(memb_row + 10, excel_col).Value = needs_interp
	objExcel.Cells(memb_row + 11, excel_col).Value = cl_alias
	objExcel.Cells(memb_row + 12, excel_col).Value = hisp_latino

	client_array = client_array & ref_num & ","
	
	'Going to the next MEMB------------------------------------------------------------------------------------------
	transmit
	
	'finding the last MEMB
	EMReadScreen enter_a_valid, 13, 24, 2
	IF enter_a_valid = "ENTER A VALID" THEN EXIT DO
	
	excel_col = excel_col + 1
LOOP

'splitting the array of ref numbers
client_array = client_array & "END"
client_array = replace(client_array, ",END", "")
client_array = split(client_array, ",")

'Dealing with the client's name...
IF maxis_enviro = "PRODU" OR maxis_enviro = "INQUI" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & vbNewLine & "You are not running this in the training region. The script will ask for member names to ensure the privacy of the real cases you are copying.", vbInformation + vbSystemModal, "Reality Detected" 

'Grabbing client names...if we are in training
REDIM client_multi_array(ubound(client_array), 3)
client_position = 0
CALL write_value_and_transmit("01", 20, 76)

IF maxis_enviro = "TRAIN" THEN 
	FOR EACH client IN client_array
		'Reading ref num
		EMReadScreen client_multi_array(client_position, 0), 2, 4, 33
		'Reading first name
		EMReadScreen client_multi_array(client_position, 1), 25, 6, 30
			client_multi_array(client_position, 1) = replace(client_multi_array(client_position, 1), "_", "")
		'reading the middle initial
		EMReadScreen client_multi_array(client_position, 2), 1, 6, 79
		'reading the last name
		EMReadScreen client_multi_array(client_position, 3), 12, 6, 63
			client_multi_array(client_position, 3) = replace(client_multi_array(client_position, 3), "_", "")
		transmit
		client_position = client_position + 1
	NEXT
END IF

'gathering/confirming the clients' names and getting APPL information	
DO
	err_msg = ""
	
	'resizing this dialog depending on the number of peeps
	dlg_height = 100 + (ubound(client_array) * 20)
	
	BeginDialog Dialog1, 0, 0, 246, dlg_height, "Name and APPL Info"
      DropListBox 60, 35, 65, 15, "Select one:"+chr(9)+"CM"+chr(9)+"CM -1"+chr(9)+"CM -2"+chr(9)+"CM -3"+chr(9)+"CM -4"+chr(9)+"CM -5", appl_month
	  EditBox 185, 30, 30, 15, appl_day	
	  Text 10, 35, 45, 10, "APPL Month:"
      Text 10, 55, 35, 10, "Ref Num"
      Text 70, 55, 65, 10, "First Name"
      Text 140, 55, 15, 10, "M.I."
      Text 170, 55, 65, 10, "Last Name"
      Text 140, 35, 35, 10, "APPL Day:"
      dlg_row = 70
	  FOR i = 0 TO ubound(client_array)
		Text 10, dlg_row, 35, 10, client_multi_array(i, 0)
		EditBox 65, dlg_row, 70, 15, client_multi_array(i, 3)
		EditBox 140, dlg_row, 20, 15, client_multi_array(i, 2)
		EditBox 170, dlg_row, 70, 15, client_multi_array(i, 1)
		dlg_row = dlg_row + 20
	  NEXT
	  ButtonGroup ButtonPressed
        OkButton 5, 10, 50, 15
        CancelButton 55, 10, 50, 15	  
	EndDialog
	
	'calling the dialog
	DIALOG
	cancel_confirmation
	'validating the data
	IF appl_month = "Select one:" THEN err_msg = err_msg & vbCr & "* Please select an APPL month."
	IF appl_day = "" OR IsNumeric(appl_day) = FALSE THEN  err_msg = err_msg & vbCr & "* Please enter a numeric APPL day."
	FOR i = 0 TO ubound(client_array)
		IF client_multi_array(i, 3) = "" THEN err_msg = err_msg & vbCr & "* Please enter a first name for member " & client_multi_array(i, 0) & "."
		IF client_multi_array(i, 2) = "" THEN err_msg = err_msg & vbCr & "* Please enter a middle initial for member " & client_multi_array(i, 0) & "."
		IF client_multi_array(i, 1) = "" THEN err_msg = err_msg & vbCr & "* Please enter a last name for member " & client_multi_array(i, 0) & "."
	NEXT
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue.", vbInformation + vbSystemModal, "Critical Error"
LOOP UNTIL ButtonPressed = -1 AND err_msg = ""

CALL check_for_MAXIS(false)

'writing the APPL information
objExcel.Cells(4, 3).Value = appl_month
objExcel.Cells(5, 3).Value = appl_day

'writing the name information to the spreadsheet
FOR i = 0 TO ubound(client_array)
	objExcel.Cells(6, i + 3).Value = client_multi_array(i, 1)	'<<< last name
	objExcel.Cells(7, i + 3).Value = client_multi_array(i, 3)	'<<< first name
	objExcel.Cells(8, i + 3).Value = client_multi_array(i, 2)	'<<< middle initial
NEXT

'Navigating to MEMI to grab that information------------------------------------------------------------------------------------------
CALL navigate_to_MAXIS_screen("STAT", "MEMI")

excel_col = 3
FOR EACH client IN client_array
	STATS_manualtime = STATS_manualtime + 11 		'<<< 11.43 seconds to write MEMI data per client
	CALL write_value_and_transmit(client, 20, 76)
	'reading MEMI for each peep
	EMReadScreen marital_status, 1, 7, 49
	EMReadScreen spouse_ref_num, 2, 8, 49
		spouse_ref_num = replace(spouse_ref_num, "_", "")
	EMReadScreen last_grade, 2, 9, 49
	EMReadScreen citizen_yn, 1, 10, 49
	
	'and writing it into the template
	objExcel.Cells(memi_row, excel_col).Value = marital_status
	objExcel.Cells(memi_row + 1, excel_col).Value = spouse_ref_num
	objExcel.Cells(memi_row + 2, excel_col).Value = last_grade
	objExcel.Cells(memi_row + 3, excel_col).Value = citizen_yn

	excel_col = excel_col + 1
NEXT

'Navigating to ADDR------------------------------------------------------------------------------------------
CALL navigate_to_MAXIS_screen("STAT", "ADDR")
STATS_manualtime = STATS_manualtime + 35		'<<< average time...if no mailing addr info...about 31 seconds...if mailing addr info...about 39 seconds
'reading ADDR information
EMReadScreen addr_line1, 22, 6, 43
	addr_line1 = replace(addr_line1, "_", "")
EMReadScreen addr_line2, 22, 7, 43
	addr_line2 = replace(addr_line2, "_", "")
EMReadScreen addr_city, 15, 8, 43
	addr_city = replace(addr_city, "_", "")
EMReadScreen addr_zip, 5, 9, 43
EMReadScreen addr_county, 2, 9, 66
EMReadScreen addr_verif, 2, 9, 74
EMReadScreen addr_homeless, 1, 10, 43
EMReadScreen addr_reserv, 1, 10, 74
EMReadScreen mail_line1, 22, 13, 43
	mail_line1 = replace(mail_line1, "_", "")
EMReadScreen mail_line2, 22, 14, 43
	mail_line2 = replace(mail_line2, "_", "")
EMReadScreen mail_city, 15, 15, 43
	mail_city = replace(mail_city, "_", "")
EMReadScreen mail_zip, 5, 16, 15
EMReadScreen phone1, 14, 17, 45
	phone1 = replace(phone1, " ) ", "-")
	phone1 = replace(phone1, " ", "-")
EMReadScreen phone2, 14, 18, 45
	phone2 = replace(phone2, " ) ", "-")
	phone2 = replace(phone2, " ", "-")
EMReadScreen phone3, 14, 19, 45
	phone3 = replace(phone3, " ) ", "-")
	phone3 = replace(phone3, " ", "-")

'and writing it into the template
objExcel.Cells(addr_row, 3).Value = addr_line1
objExcel.Cells(addr_row + 1, 3).Value = addr_line2
objExcel.Cells(addr_row + 2, 3).Value = addr_city
objExcel.Cells(addr_row + 3, 3).Value = addr_zip
objExcel.Cells(addr_row + 4, 3).Value = addr_county
objExcel.Cells(addr_row + 5, 3).Value = addr_verif
objExcel.Cells(addr_row + 6, 3).Value = addr_homeless
objExcel.Cells(addr_row + 7, 3).Value = addr_reserv
objExcel.Cells(addr_row + 8, 3).Value = mail_line1
objExcel.Cells(addr_row + 9, 3).Value = mail_line2
objExcel.Cells(addr_row + 10, 3).Value = mail_city
objExcel.Cells(addr_row + 11, 3).Value = mail_zip
IF phone1 <> "___-___-____" THEN objExcel.Cells(addr_row + 12, 3).Value = phone1
IF phone2 <> "___-___-____" THEN objExcel.Cells(addr_row + 13, 3).Value = phone2
IF phone3 <> "___-___-____" THEN objExcel.Cells(addr_row + 14, 3).Value = phone3

'Navigating to TYPE------------------------------------------------------------------------------------------
CALL navigate_to_MAXIS_screen("STAT", "TYPE")

excel_col = 3
type_row = 6
'determining which programs are active...necessary for grabbing income information later on
cash_case = FALSE
health_care_case = FALSE
snap_case = FALSE
FOR EACH client IN client_array
	'adding to STATS_manualtime -------------------
	STATS_manualtime = STATS_manualtime + 11
	
	EMReadScreen cash_appl, 1, type_row, 28
	IF cash_appl = "Y" THEN cash_case = TRUE
	EMReadScreen hc_appl, 1, type_row, 37
	IF hc_appl = "Y" THEN health_care_case = TRUE
	EMReadScreen snap_appl, 1, type_row, 46
	IF snap_appl = "Y" THEN snap_case = TRUE
	
	objExcel.Cells(38, excel_col).Value = cash_appl
	objExcel.Cells(39, excel_col).Value = hc_appl
	objExcel.Cells(40, excel_col).Value = snap_appl
	
	excel_col = excel_col + 1
	type_row = type_row + 1
NEXT

'Navigating to PROG------------------------------------------------------------------------------------------
CALL navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen migrant_farmworker, 1, 18, 67

excel_col = 3
FOR EACH client IN client_array
	'adding to STATS_manualtime -------------------
	STATS_manualtime = STATS_manualtime + 4
	objExcel.Cells(41, excel_col).Value = migrant_farmworker
	excel_col = excel_col + 1
NEXT

'Navigating to REVW------------------------------------------------------------------------------------------
IF health_care_case = TRUE THEN 
	CALL navigate_to_MAXIS_screen("STAT", "REVW")
	'adding to STATS_manualtime -----------------------
	STATS_manualtime = STATS_manualtime + 12
	
	CALL write_value_and_transmit("X", 5, 71)
	EMReadScreen ir_date, 8, 8, 27
	EMReadScreen ir_ar_date, 8, 8, 71
	EMReadScreen exempt_ir_ar, 1, 9, 71
	
	IF ir_date <> "__ 01 __" THEN objExcel.Cells(revw_row, 3).Value = replace(ir_date, " ", "/")
	IF ir_ar_date <> "__ 01 __" THEN objExcel.Cells(revw_row, 3).Value = replace(ir_ar_date, " ", "/")
	objExcel.Cells(revw_row + 1, 3).Value = exempt_ir_ar	
	transmit
END IF

'Now going through each panel------------------------------------------------------------------------------------------
'ABPS
CALL navigate_to_MAXIS_screen("STAT", "ABPS")
EMReadScreen num_of_abps, 1, 2, 78
IF num_of_abps <> "0" THEN 
	'adding to STATS_manualtime -------------------
	STATS_manualtime = STATS_manualtime + 7
	
	EMReadScreen abps_support_coop, 1, 4, 73
	EMReadScreen abps_good_cause, 1, 5, 47

	objExcel.Cells(abps_row, 3).Value = abps_support_coop
	objExcel.Cells(abps_row + 1, 3).Value = abps_good_cause
END IF

'ACCT
CALL navigate_to_MAXIS_screen("STAT", "ACCT")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_acct, 1, 2, 78
	IF num_of_acct <> "0" THEN 
		'adding to STATS_manualtime ------------------------------------
		STATS_manualtime = STATS_manualtime + 47
		
		'reading
		EMReadScreen acct_account_type, 2, 6, 44
		EMReadScreen acct_account_num, 20, 7, 44
			acct_account_num = replace(acct_account_num, "_", "")
		EMReadScreen acct_account_loc, 20, 8, 44
			acct_account_loc = replace(acct_account_loc, "_", "")
		EMReadScreen acct_account_bal, 8, 10, 46
			acct_account_bal = trim(acct_account_bal)
		EMReadScreen acct_bal_ver_code, 1, 10, 64
		EMReadScreen acct_account_bal_dt, 8, 11, 44
			acct_account_bal_dt = replace(acct_account_bal_dt, " ", "/")
		EMReadScreen acct_withdraw_pen, 8, 12, 46
			acct_withdraw_pen = replace(acct_withdraw_pen, "_", "")
			acct_withdraw_pen = trim(acct_withdraw_pen)
		EMReadScreen acct_cash_count, 1, 14, 50
		EMReadScreen acct_snap_count, 1, 14, 57
		EMReadScreen acct_hc_count, 1, 14, 64
		EMReadScreen acct_grh_count, 1, 14, 71
		EMReadScreen acct_ive_count, 1, 14, 78
		EMReadScreen acct_joint_owner, 1, 15, 44
		EMReadScreen acct_share_ratio, 5, 15, 76
		EMReadScreen acct_interest_mo, 2, 17, 57
		EMReadScreen acct_interest_yr, 2, 17, 60
		
		'writing
		CALL check_for_data_validation(acct_row, excel_col, acct_account_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(acct_row + 1, excel_col).Value = acct_account_num
		objExcel.Cells(acct_row + 2, excel_col).Value = acct_account_loc
		objExcel.Cells(acct_row + 3, excel_col).Value = acct_account_bal
		CALL check_for_data_validation(acct_row +  4, excel_col, acct_bal_ver_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(acct_row + 5, excel_col).Value = acct_account_bal_dt
		objExcel.Cells(acct_row + 6, excel_col).Value = acct_withdraw_pen
		CALL check_for_data_validation(acct_row +  7, excel_col, acct_cash_count, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(acct_row +  8, excel_col, acct_snap_count, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(acct_row +  9, excel_col, acct_hc_count, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(acct_row +  10, excel_col, acct_grh_count, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(acct_row +  11, excel_col, acct_ive_count, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(acct_row +  12, excel_col, acct_joint_owner, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(acct_row + 13, excel_col).Value = acct_share_ratio
		objExcel.Cells(acct_row + 14, excel_col).Value = acct_interest_mo
		objExcel.Cells(acct_row + 15, excel_col).Value = acct_interest_yr	
	END IF
	excel_col = excel_col + 1
NEXT

'ACUT
CALL navigate_to_MAXIS_screen("STAT", "ACUT")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_acut, 1, 2, 78
	IF num_of_acut <> "0" THEN 
		'Adding to STATS_manualtime...............
		STATS_manualtime = STATS_manualtime + 24
		
		'reading
		EMReadScreen acut_shared, 1, 6, 42
		EMReadScreen acut_heat, 8, 10, 61
			acut_heat = trim(replace(acut_heat, "_", ""))
		EMReadScreen acut_air, 8, 11, 61
			acut_air = trim(replace(acut_air, "_", ""))
		EMReadScreen acut_electric, 8, 12, 61
			acut_electric = trim(replace(acut_electric, "_", ""))
		EMReadScreen acut_fuel, 8, 13, 61
			acut_fuel = trim(replace(acut_fuel, "_", ""))
		EMReadScreen acut_garbage, 8, 14, 61
			acut_garbage = trim(replace(acut_garbage, "_", ""))
		EMReadScreen acut_water, 8, 15, 61
			acut_water = trim(replace(acut_water, "_", ""))
		EMReadScreen acut_sewer, 8, 16, 61
			acut_sewer = trim(replace(acut_sewer, "_", ""))
		EMReadScreen acut_other, 8, 17, 61
			acut_other = trim(replace(acut_other, "_", ""))
		EMReadScreen acut_heat_verif, 1, 10, 55
		EMReadScreen acut_air_verif, 1, 11, 55
		EMReadScreen acut_electric_verif, 1, 12, 55
		EMReadScreen acut_fuel_verif, 1, 13, 55
		EMReadScreen acut_garbage_verif, 1, 14, 55
		EMReadScreen acut_water_verif, 1, 15, 55
		EMReadScreen acut_sewer_verif, 1, 16, 55
		EMReadScreen acut_other_verif, 1, 17, 55
		EMReadScreen acut_phone, 1, 18, 55
		
		'writing
		IF acut_shared <> "_" 			THEN CALL check_for_data_validation(acut_row, excel_col, acut_shared, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_heat <> "" 				THEN objExcel.Cells(acut_row + 1, excel_col).Value = acut_heat
		IF acut_heat_verif <> "_" 		THEN CALL check_for_data_validation(acut_row + 2, excel_col, acut_heat_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_air <> "" 				THEN objExcel.Cells(acut_row + 3, excel_col).Value = acut_air
		IF acut_air_verif <> "_" 		THEN CALL check_for_data_validation(acut_row + 4, excel_col, acut_air_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_electric <> "" 			THEN objExcel.Cells(acut_row + 5, excel_col).Value = acut_electric
		IF acut_electric_verif <> "_" 	THEN CALL check_for_data_validation(acut_row + 6, excel_col, acut_electric_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_fuel <> "" 				THEN objExcel.Cells(acut_row + 7, excel_col).Value = acut_fuel
		IF acut_fuel_verif <> "_" 		THEN CALL check_for_data_validation(acut_row + 8, excel_col, acut_fuel_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_garbage <> "" 			THEN objExcel.Cells(acut_row + 9, excel_col).Value = acut_garbage
		IF acut_garbage_verif <> "_" 	THEN CALL check_for_data_validation(acut_row + 10, excel_col, acut_garbage_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_water <> "" 			THEN objExcel.Cells(acut_row + 11, excel_col).Value = acut_water
		IF acut_water_verif <> "_" 		THEN CALL check_for_data_validation(acut_row + 12, excel_col, acut_water_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_sewer <> "" 			THEN objExcel.Cells(acut_row + 13, excel_col).Value = acut_sewer
		IF acut_sewer_verif <> "_" 		THEN CALL check_for_data_validation(acut_row + 14, excel_col, acut_sewer_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF acut_other <> "" 			THEN objExcel.Cells(acut_row + 15, excel_col).Value = acut_other
		IF acut_other_verif <> "_" 		THEN CALL check_for_data_validation(acut_row + 16, excel_col, acut_other_verif, objExcel, objWorkbook, objTemplate, objNewSheet)		
		IF acut_phone <> "" 			THEN CALL check_for_data_validation(acut_row + 17, excel_col, acut_phone, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'BILS
CALL navigate_to_MAXIS_screen("STAT", "BILS")
EMReadScreen num_of_bils, 1, 2, 78
IF num_of_bils <> "0" THEN 
	'creating m-d array so I don't have to write this 9 times -- Robert
	REDIM bils_array (8, 6)
	'(i, 0) = ref num 
	'(i, 1) = date
	'(i, 2) = service 
	'(i, 3) = gross amount 
	'(i, 4) = third party payments
	'(i, 5) = verif
	'(i, 6) = bill type
	
	'reading
	bils_read_row = 6
	FOR i = 0 TO 8
		EMReadScreen bils_array(i, 0), 2, bils_read_row + i, 26
		EMReadScreen bils_array(i, 1), 8, bils_read_row + i, 30
		EMReadScreen bils_array(i, 2), 2, bils_read_row + i, 40
		EMReadScreen bils_array(i, 3), 9, bils_read_row + i, 45
			bils_array(i, 3) = trim(replace(bils_array(i, 3), "_", ""))
		EMReadScreen bils_array(i, 4), 9, bils_read_row + i, 57
			bils_array(i, 4) = trim(replace(bils_array(i, 4), "_", ""))
		EMReadScreen bils_array(i, 5), 2, bils_read_row + i, 67
		EMReadScreen bils_array(i, 6), 1, bils_read_row + i, 71	
	NEXT
	
	'writing
	FOR i = 0 TO 8
		IF bils_array(i, 0) <> "__" THEN 
			'adding to the manual time for each row that the script is copying to Excel...
			STATS_manualtime = STATS_manualtime + 25
			
			objExcel.Cells(bils_row, 3).Value = bils_array(i, 0)
				bils_row = bils_row + 1
			IF bils_array(i, 1) <> "__ __ __" 		THEN objExcel.Cells(bils_row, 3).Value = replace(bils_array(i, 1), " ", "/")
				bils_row = bils_row + 1
			IF bils_array(i, 2) <> "__" 			THEN CALL check_for_data_validation(bils_row, 3, bils_array(i, 2), objExcel, objWorkbook, objTemplate, objNewSheet)
				bils_row = bils_row + 1
			IF bils_array(i, 3) <> ""				THEN objExcel.Cells(bils_row, 3).Value = bils_array(i, 3)
				bils_row = bils_row + 1
			IF bils_array(i, 4) <> ""				THEN objExcel.Cells(bils_row, 3).Value = bils_array(i, 4)
				bils_row = bils_row + 1
			IF bils_array(i, 5) <> "__" 			THEN CALL check_for_data_validation(bils_row, 3, bils_array(i, 5), objExcel, objWorkbook, objTemplate, objNewSheet)
				bils_row = bils_row + 1
			IF bils_array(i, 6) <> "_" 				THEN objExcel.Cells(bils_row, 3).Value = bils_array(i, 6)
				bils_row = bils_row + 1
		END IF
	NEXT
END IF

'BUSI
CALL navigate_to_MAXIS_screen("STAT", "BUSI")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_busi, 1, 2, 78
	IF num_of_busi <> "0" THEN 
		'looking for the first active BUSI panel
		DO 
			EMReadScreen panel_end_date, 8, 5, 72
			IF panel_end_date = "__ __ __" THEN 
				'adding to STATS_manualtime --------------------------
				STATS_manualtime = STATS_manualtime + 55
				'the stats manual time on BUSI is a bit variable given the number of data points... this run time assumes cash and SNAP
			
				'reading
				'basic info
				EMReadScreen busi_type, 2, 5, 37
				EMReadScreen busi_start_date, 8, 5, 55
				'not reading busi end date because the script has already skipped over that
				
				'going in to the income calc pop up
				CALL write_value_and_transmit("X", 6, 26)
				'income
				'cash
				EMReadScreen busi_cash_inc_retro, 8, 9, 43
					busi_cash_inc_retro = trim(replace(busi_cash_inc_retro, "_", ""))
				EMReadScreen busi_cash_inc_prosp, 8, 9, 59
					busi_cash_inc_prosp = trim(replace(busi_cash_inc_prosp, "_", ""))
				EMReadScreen busi_cash_inc_verif, 1, 9, 73
				'ive
				EMReadScreen busi_ive_inc_prosp, 8, 10, 59
					busi_ive_inc_prosp = trim(replace(busi_ive_inc_prosp, "_", ""))
				EMReadScreen busi_ive_inc_verif, 1, 10, 73
				'snap
				EMReadScreen busi_snap_inc_retro, 8, 11, 43
					busi_snap_inc_retro = trim(replace(busi_snap_inc_retro, "_", ""))
				EMReadScreen busi_snap_inc_prosp, 8, 11, 59
					busi_snap_inc_prosp = trim(replace(busi_snap_inc_prosp, "_", ""))
				EMReadScreen busi_snap_inc_verif, 1, 11, 73
				'hc
				EMReadScreen busi_hc_a_inc_prosp, 8, 12, 59
					busi_hc_a_inc_prosp = trim(replace(busi_hc_a_inc_prosp, "_", ""))
				EMReadScreen busi_hc_a_inc_verif, 1, 12, 73
				EMReadScreen busi_hc_b_inc_prosp, 8, 13, 59
					busi_hc_b_inc_prosp = trim(replace(busi_hc_b_inc_prosp, "_", ""))
				EMReadScreen busi_hc_b_inc_verif, 1, 13, 73
				'expenses
				'cash	
				EMReadScreen busi_cash_exp_retro, 8, 15, 43
					busi_cash_exp_retro = trim(replace(busi_cash_exp_retro, "_", ""))
				EMReadScreen busi_cash_exp_prosp, 8, 15, 59
					busi_cash_exp_retro = trim(replace(busi_cash_exp_prosp, "_", ""))
				EMReadScreen busi_cash_exp_verif, 1, 15, 73
				'ive
				EMReadScreen busi_ive_exp_prosp, 8, 16, 59
					busi_ive_exp_prosp = trim(replace(busi_ive_exp_prosp, "_", ""))
				EMReadScreen busi_ive_exp_verif, 1, 16, 73
				'snap
				EMReadScreen busi_snap_exp_retro, 8, 17, 43
					busi_snap_exp_retro = trim(replace(busi_snap_exp_retro, "_", ""))
				EMReadScreen busi_snap_exp_prosp, 8, 17, 59
					busi_snap_exp_prosp = trim(replace(busi_snap_exp_prosp, "_", ""))
				EMReadScreen busi_snap_exp_verif, 1, 17, 73
				'hc
				EMReadScreen busi_hc_a_exp_prosp, 8, 18, 59
					busi_hc_a_exp_prosp = trim(replace(busi_hc_a_exp_prosp, "_", ""))
				EMReadScreen busi_hc_a_exp_verif, 1, 18, 73
				EMReadScreen busi_hc_b_exp_prosp, 8, 19, 59
					busi_hc_b_exp_prosp = trim(replace(busi_hc_b_exp_prosp, "_", ""))
				EMReadScreen busi_hc_b_exp_verif, 1, 19, 73
				'leaving pop up
				PF3
				
				EMReadScreen busi_self_emp_retro_hours, 3, 13, 60
					busi_self_emp_retro_hours = trim(replace(busi_self_emp_retro_hours, "_", ""))
				EMReadScreen busi_self_emp_prosp_hours, 3, 13, 74
					busi_self_emp_prosp_hours = trim(replace(busi_self_emp_prosp_hours, "_", ""))
					
				'going to HC inc/exp pop up
				CALL write_value_and_transmit("X", 17, 27)
				EMReadScreen busi_hc_ttl_inc_a, 8, 7, 54
					busi_hc_ttl_inc_a = trim(replace(busi_hc_ttl_inc_a, "_", ""))
				EMReadScreen busi_hc_ttl_inc_b, 8, 8, 54
					busi_hc_ttl_inc_b = trim(replace(busi_hc_ttl_inc_b, "_", ""))
				EMReadScreen busi_hc_ttl_exp_a, 8, 11, 54
					busi_hc_ttl_exp_a = trim(replace(busi_hc_ttl_exp_a, "_", ""))
				EMReadScreen busi_hc_ttl_exp_b, 8, 12, 54
					busi_hc_ttl_exp_b = trim(replace(busi_hc_ttl_exp_b, "_", ""))
				EMReadScreen busi_hc_ttl_hours, 3, 18, 58
					busi_hc_ttl_hours = trim(replace(busi_hc_ttl_hours, "_", ""))
				'leaving pop up
				PF3
		
				'writing
				CALL check_for_data_validation(busi_row, excel_col, busi_type, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 1, excel_col).Value = replace(busi_start_date, " ", "/")
				'busi_row + 2 is busi end date...omitted because the script is skipping BUSI panels with an end date
				objExcel.Cells(busi_row + 3, excel_col).Value = busi_cash_inc_retro
				objExcel.Cells(busi_row + 4, excel_col).Value = busi_cash_inc_prosp
				IF busi_cash_inc_verif <> "_" THEN CALL check_for_data_validation(busi_row + 5, excel_col, busi_cash_inc_verif, objExcel, objWorkbook, objTemplate, objNewSheet) ' <<< problem, the BUSI panel data validation is a different format...it is hard coded rather than a list in controls
				objExcel.Cells(busi_row + 6, excel_col).Value = busi_ive_inc_prosp
				IF busi_ive_exp_verif <> "_" THEN CALL check_for_data_validation(busi_row + 7, excel_col, busi_ive_exp_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 8, excel_col).Value = busi_snap_inc_retro
				objExcel.Cells(busi_row + 9, excel_col).Value = busi_snap_inc_prosp
				IF busi_snap_inc_verif <> "_" THEN CALL check_for_data_validation(busi_row + 10, excel_col, busi_snap_inc_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 11, excel_col).Value = busi_hc_a_inc_prosp
				IF busi_hc_a_inc_verif <> "_" THEN CALL check_for_data_validation(busi_row + 12, excel_col, busi_hc_a_inc_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 13, excel_col).Value = busi_hc_b_inc_prosp
				IF busi_hc_b_inc_verif <> "_" THEN CALL check_for_data_validation(busi_row + 14, excel_col, busi_hc_b_inc_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 15, excel_col).Value = busi_cash_exp_retro
				objExcel.Cells(busi_row + 16, excel_col).Value = busi_cash_exp_prosp
				IF busi_cash_exp_verif <> "_" THEN CALL check_for_data_validation(busi_row + 17, excel_col, busi_cash_exp_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 18, excel_col).Value = busi_ive_exp_prosp
				IF busi_ive_exp_verif <> "_" THEN CALL check_for_data_validation(busi_row + 19, excel_col, busi_ive_exp_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 20, excel_col).Value = busi_snap_exp_retro
				objExcel.Cells(busi_row + 21, excel_col).Value = busi_snap_exp_prosp
				IF busi_snap_exp_verif <> "_" THEN CALL check_for_data_validation(busi_row + 22, excel_col, busi_snap_exp_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 23, excel_col).Value = busi_hc_a_exp_prosp
				IF busi_hc_a_exp_verif <> "_" THEN CALL check_for_data_validation(busi_row + 24, excel_col, busi_hc_a_exp_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 25, excel_col).Value = busi_hc_b_exp_prosp
				IF busi_hc_b_exp_verif <> "_" THEN CALL check_for_data_validation(busi_row + 26, excel_col, busi_hc_b_exp_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(busi_row + 27, excel_col).Value = busi_self_emp_prosp_hours
				objExcel.Cells(busi_row + 28, excel_col).Value = busi_self_emp_retro_hours
				objExcel.Cells(busi_row + 29, excel_col).Value = busi_hc_ttl_inc_a
				objExcel.Cells(busi_row + 30, excel_col).Value = busi_hc_ttl_inc_b
				objExcel.Cells(busi_row + 31, excel_col).Value = busi_hc_ttl_exp_a
				objExcel.Cells(busi_row + 32, excel_col).Value = busi_hc_ttl_exp_b
				objExcel.Cells(busi_row + 33, excel_col).Value = busi_hc_ttl_hours
				
				'and leave b/c we have found one for this client and that is all the template can handle
				EXIT DO
			ELSE
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
				IF enter_a_valid = "ENTER A VALID" THEN EXIT DO
			END IF
		LOOP
	END IF
	excel_col = excel_col + 1
NEXT

'CARS
CALL navigate_to_MAXIS_screen("STAT", "CARS")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_cars, 1, 2, 78
	IF num_of_cars <> "0" THEN 
		'adding to STATS_manualtime ----------------------
		STATS_manualtime = STATS_manualtime + 50
			
		
		'reading
		EMReadScreen cars_type, 1, 6, 43
		EMReadScreen cars_year, 4, 8, 31
			cars_year = replace(cars_year, "_", "")
		EMReadScreen cars_make, 15, 8, 43
			cars_make = trim(cars_make)
			cars_make = replace(cars_make, "_", "")
		EMReadScreen cars_model, 15, 8, 66
			cars_model = trim(cars_model)
			cars_model = replace(cars_model, "_", "")
		EMReadScreen cars_trade_in, 8, 9, 45
			cars_trade_in = trim(cars_trade_in)
			cars_trade_in = replace(cars_trade_in, "_", "")
		EMReadScreen cars_loan_val, 8, 9, 62
			cars_loan_val = trim(cars_loan_val)
			cars_loan_val = replace(cars_loan_val, "_", "")
		EMReadScreen cars_source_code, 1, 9, 80
			cars_source_code = replace(cars_source_code, "_", "")
		EMReadScreen cars_own_verif, 1, 10, 60
			cars_own_verif = replace(cars_own_verif, "_", "")
		EMReadScreen cars_amt_owed, 8, 12, 45
			cars_amt_owed = trim(cars_amt_owed)
			cars_amt_owed = replace(cars_amt_owed, "_", "")
		EMReadScreen cars_amt_owed_ver, 1, 12, 60
		EMReadScreen cars_as_of_date, 8, 13, 43
		EMReadScreen cars_use, 1, 15, 43
		EMReadScreen cars_hc_cl_bene, 1, 15, 76
		EMReadScreen cars_joint_own, 1, 16, 43
		EMReadScreen cars_share_ratio, 5, 16, 76
		
		'writing
		CALL check_for_data_validation(cars_row, excel_col, cars_type, objExcel, objWorkbook, objTemplate, objNewSheet)	
		IF cars_year <> "" 					THEN objExcel.Cells(cars_row + 1, excel_col).Value = cars_year
		IF cars_make <> "" 					THEN objExcel.Cells(cars_row + 2, excel_col).Value = cars_make
		IF cars_model <> "" 				THEN objExcel.Cells(cars_row + 3, excel_col).Value = cars_model
		IF cars_trade_in <> "" 				THEN objExcel.Cells(cars_row + 4, excel_col).Value = cars_trade_in
		IF cars_loan_val <> "" 				THEN objExcel.Cells(cars_row + 5, excel_col).Value = cars_loan_val
		IF cars_source_code <> "" 			THEN CALL check_for_data_validation(cars_row + 6, excel_col, cars_source_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF cars_own_verif <> "" 			THEN CALL check_for_data_validation(cars_row + 7, excel_col, cars_own_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF cars_amt_owed <> "" 				THEN objExcel.Cells(cars_row + 8, excel_col).Value = cars_amt_owed
		IF cars_amt_owed_ver <> "_" 		THEN CALL check_for_data_validation(cars_row + 9, excel_col, cars_amt_owed_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF cars_as_of_date <> "__ __ __" 	THEN objExcel.Cells(cars_row + 10, excel_col).Value = replace(cars_as_of_date, " ", "/")
		IF cars_use <> "_" 					THEN CALL check_for_data_validation(cars_row + 11, excel_col, cars_use, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(cars_row + 12, excel_col, cars_hc_cl_bene, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF cars_joint_own <> "_" 			THEN CALL check_for_data_validation(cars_row + 13, excel_col, cars_joint_own, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(cars_row + 14, excel_col).Value = cars_share_ratio
	END IF
	excel_col = excel_col + 1
NEXT

'CASH
CALL navigate_to_MAXIS_screen("STAT", "CASH")
excel_col = 3
FOR EACH client IN client_array
	'adding to STATS_manualtime --------------------------------
	STATS_manualtime = STATS_manualtime + 7
	'not a lot going on

	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen cash_amt, 8, 8, 39
	cash_cash_amt = trim(cash_amt)
	cash_cash_amt = replace(cash_amt, "_", "")
	
	'only value on CASH
	objExcel.Cells(cash_row, excel_col).Value = cash_cash_amt	
	
	excel_col = excel_col + 1
NEXT

'COEX
CALL navigate_to_MAXIS_screen("STAT", "COEX")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_coex, 1, 2, 78
	IF num_of_coex <> "0" THEN
		'adding to STATS_manualtime --------------------------
		STATS_manualtime = STATS_manualtime + 30		
		
		'reading the COEX bits
		EMReadScreen coex_retro_support, 8, 10, 45
			coex_retro_support = trim(replace(coex_retro_support, "_", ""))
		EMReadScreen coex_prosp_support, 8, 10, 63
			coex_prosp_support = trim(replace(coex_prosp_support, "_", ""))
		EMReadScreen coex_support_verif, 1, 10, 36
		EMReadScreen coex_retro_alimony, 8, 11, 45
			coex_retro_alimony = trim(replace(coex_retro_alimony, "_", ""))
		EMReadScreen coex_prosp_alimony, 8, 11, 63
			coex_prosp_alimony = trim(replace(coex_prosp_alimony, "_", ""))
		EMReadScreen coex_alimony_verif, 1, 11, 36
		EMReadScreen coex_retro_tax_dep, 8, 12, 45
			coex_retro_tax_dep = trim(replace(coex_retro_tax_dep, "_", ""))
		EMReadScreen coex_prosp_tax_dep, 8, 12, 63
			coex_prosp_tax_dep = trim(replace(coex_prosp_tax_dep, "_", ""))
		EMReadScreen coex_tax_dep_verif, 1, 12, 36
		EMReadScreen coex_retro_other, 8, 13, 45
			coex_retro_other = trim(replace(coex_retro_other, "_", ""))
		EMReadScreen coex_prosp_other, 8, 13, 63
			coex_prosp_other = trim(replace(coex_prosp_other, "_", ""))
		EMReadScreen coex_other_verif, 1, 13, 45
		EMReadScreen coex_change_in_circ, 1, 17, 61
		'going in to the hc pop up
		CALL write_value_and_transmit("X", 18, 44)
		EMReadScreen coex_hc_support, 8, 6, 38
			coex_hc_support = trim(replace(coex_hc_support, "_", ""))
		EMReadScreen coex_hc_alimony, 8, 7, 38
			coex_hc_alimony = trim(replace(coex_hc_alimony, "_", ""))
		EMReadScreen coex_hc_tax_dep, 8, 8, 38
			coex_hc_tax_dep = trim(replace(coex_hc_tax_dep, "_", ""))
		EMReadScreen coex_hc_other, 8, 8, 38
			coex_hc_other = trim(replace(coex_hc_other, "_", ""))
		PF3
		
		'writing to template
		IF coex_retro_support <> "" 	THEN objExcel.Cells(coex_row, excel_col).Value = coex_retro_support
		IF coex_prosp_support <> "" 	THEN objExcel.Cells(coex_row + 1, excel_col).Value = coex_prosp_support
		IF coex_support_verif <> "_" 	THEN CALL check_for_data_validation(coex_row + 2, excel_col, coex_support_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF coex_retro_alimony <> "" 	THEN objExcel.Cells(coex_row + 3, excel_col).Value = coex_retro_alimony
		IF coex_prosp_alimony <> "" 	THEN objExcel.Cells(coex_row + 4, excel_col).Value = coex_prosp_alimony
		IF coex_alimony_verif <> "_" 	THEN CALL check_for_data_validation(coex_row + 5, excel_col, coex_alimony_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF coex_retro_tax_dep <> "" 	THEN objExcel.Cells(coex_row + 6, excel_col).Value = coex_retro_tax_dep
		IF coex_prosp_tax_dep <> "" 	THEN objExcel.Cells(coex_row + 7, excel_col).Value = coex_prosp_tax_dep
		IF coex_tax_dep_verif <> "_" 	THEN CALL check_for_data_validation(coex_row + 8, excel_col, coex_tax_dep_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF coex_retro_other <> "" 		THEN objExcel.Cells(coex_row + 9, excel_col).Value = coex_retro_other
		IF coex_prosp_other <> "" 		THEN objExcel.Cells(coex_row + 10, excel_col).Value = coex_prosp_other
		IF coex_other_verif <> "_" 		THEN CALL check_for_data_validation(coex_row + 11, excel_col, coex_other_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF coex_change_in_circ <> "_" 	THEN CALL check_for_data_validation(coex_row + 12, excel_col, coex_change_in_circ, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF coex_hc_support <> "" 		THEN objExcel.Cells(coex_row + 13, excel_col).Value = coex_hc_support
		IF coex_hc_alimony <> "" 		THEN objExcel.Cells(coex_row + 14, excel_col).Value = coex_hc_alimony
		IF coex_hc_tax_dep <> "" 		THEN objExcel.Cells(coex_row + 15, excel_col).Value = coex_hc_tax_dep
		IF coex_hc_other <> "" 			THEN objExcel.Cells(coex_row + 16, excel_col).Value = coex_hc_other
	END IF
	excel_col = excel_col + 1
NEXT

'DCEX
CALL navigate_to_MAXIS_screen("STAT", "DCEX")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_dcex, 1, 2, 78
	IF num_of_dcex <> "0" THEN 
		EMReadScreen dcex_provider, 25, 6, 47
			dcex_provider = trim(replace(dcex_provider, "_", ""))
		EMReadScreen dcex_reason, 1, 7, 44
		EMReadScreen dcex_subsidy, 1, 8, 44
		'Creating a neat little array to simplify reading for 6 kids...
		'...because I'm getting too old for this ish
		REDIM dcex_array(5, 3)
		FOR i = 0 TO 5
			'(i, 0) = ref num
			'(i, 1) = verif
			'(i, 2) = retro
			'(i, 3) = prosp
			EMReadScreen dcex_array(i, 0), 2, i + 11, 29			
			EMReadScreen dcex_array(i, 1), 8, i + 11, 48
			EMReadScreen dcex_array(i, 2), 8, i + 11, 63
			EMReadScreen dcex_array(i, 3), 1, i + 11, 41
		NEXT		
	
		'writing
		objExcel.Cells(dcex_row, excel_col).Value = dcex_provider
		IF dcex_reason <> "_" THEN CALL check_for_data_validation(dcex_row + 1, excel_col, dcex_reason, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF dcex_subsidy <> "_" THEN CALL check_for_data_validation(dcex_row + 2, excel_col, dcex_subsidy, objExcel, objWorkbook, objTemplate, objNewSheet)
		placeholder_dcex_row = 3
		FOR i = 0 TO 5
			IF dcex_array(i, 0) <> "__" THEN 
				'adding to STATS_manualtime ONLY when the row is being written to Excel ---------------------
				STATS_manualtime = STATS_manualtime + 22		
				
				objExcel.Cells(dcex_row + placeholder_dcex_row, excel_col).Value = dcex_array(i, 0)
					placeholder_dcex_row = placeholder_dcex_row + 1
				objExcel.Cells(dcex_row + placeholder_dcex_row, excel_col).Value = dcex_array(i, 1)
					placeholder_dcex_row = placeholder_dcex_row + 1
				objExcel.Cells(dcex_row + placeholder_dcex_row, excel_col).Value = dcex_array(i, 2)
					placeholder_dcex_row = placeholder_dcex_row + 1
				CALL check_for_data_validation(dcex_row + placeholder_dcex_row, excel_col, dcex_array(i, 3), objExcel, objWorkbook, objTemplate, objNewSheet)
					placeholder_dcex_row = placeholder_dcex_row + 1
			END IF		
		NEXT
		'Reseting the array
		FOR i = 0 TO 5
			FOR j = 0 TO 3
				dcex_array(i, j) = ""
			NEXT
		NEXT
	END IF
	excel_col = excel_col + 1
NEXT

'DFLN
CALL navigate_to_MAXIS_screen("STAT", "DFLN")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_dfln, 1, 2, 78
	IF num_of_dfln <> "0" THEN 
		'adding to STATS_manualtime ------------------------
		STATS_manualtime = STATS_manualtime + 27
		
		'reading
		EMReadScreen dfln_conv_dt1, 8, 6, 27
		EMReadScreen dfln_conv_juris1, 30, 6, 41
			dfln_conv_juris1 = trim(dfln_conv_juris1)
			dfln_conv_juris1 = replace(dfln_conv_juris1, "_", "")
		EMReadScreen dfln_conv_state1, 2, 6, 75
		EMReadScreen dfln_conv_dt2, 8, 7, 27		
		EMReadScreen dfln_conv_juris2, 30, 7, 41
			dfln_conv_juris2 = trim(dfln_conv_juris2)
			dfln_conv_juris2 = replace(dfln_conv_juris2, "_", "")
		EMReadScreen dfln_conv_state2, 2, 6, 75
		EMReadScreen dfln_test_dt1, 8, 14, 27
		EMReadScreen dfln_tester_name1, 30, 14, 41
			dfln_tester_name1 = trim(dfln_tester_name1)
			dfln_tester_name1 = replace(dfln_tester_name1, "_", "")
		EMReadScreen dfln_result1, 2, 14, 75
		EMReadScreen dfln_test_dt2, 8, 15, 27
		EMReadScreen dfln_tester_name2, 30, 15, 41
			dfln_tester_name2 = trim(dfln_tester_name2)
			dfln_tester_name2 = replace(dfln_tester_name2, "_", "")
		EMReadScreen dfln_result2, 2, 15, 75	
	
		'writing
		IF dfln_conv_dt1 <> "__ __ __"			THEN objExcel.Cells(dfln_row, excel_col).Value = replace(dfln_conv_dt1, " ", "/")
		IF dfln_conv_juris1 <> ""				THEN objExcel.Cells(dfln_row + 1, excel_col).Value = dfln_conv_juris1
		IF dfln_conv_state1 <> "__"				THEN CALL check_for_data_validation(dfln_row + 2, excel_col, dfln_conv_state1, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF dfln_conv_dt2 <> "__ __ __"			THEN objExcel.Cells(dfln_row + 3, excel_col).Value = replace(dfln_conv_dt2, " ", "/")
		IF dfln_conv_juris2 <> ""				THEN objExcel.Cells(dfln_row + 4, excel_col).Value = dfln_conv_juris2
		IF dfln_conv_state2 <> "__"				THEN CALL check_for_data_validation(dfln_row + 5, excel_col, dfln_conv_state2, objExcel, objWorkbook, objTemplate, objNewSheet)		
		IF dfln_test_dt1 <> "__ __ __"			THEN objExcel.Cells(dfln_row + 6, excel_col).Value = replace(dfln_test_dt1, " ", "/")
		IF dfln_tester_name1 <> "" 				THEN objExcel.Cells(dfln_row + 7, excel_col).Value = dfln_tester_name1
		IF dfln_result1	<> "__" 				THEN CALL check_for_data_validation(dfln_row + 8, excel_col, dfln_result1, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF dfln_test_dt2 <> "__ __ __"			THEN objExcel.Cells(dfln_row + 9, excel_col).Value = replace(dfln_test_dt2, " ", "/")
		IF dfln_tester_name2 <> "" 				THEN objExcel.Cells(dfln_row + 10, excel_col).Value = dfln_tester_name2
		IF dfln_result2	<> "__" 				THEN CALL check_for_data_validation(dfln_row + 11, excel_col, dfln_result2, objExcel, objWorkbook, objTemplate, objNewSheet)		
	END IF	
	excel_col = excel_col + 1
NEXT

'DIET
CALL navigate_to_MAXIS_screen("STAT", "DIET")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_diet, 1, 2, 78
	IF num_of_diet <> "0" THEN 
		'adding to STATS_manualtime -------------------
		STATS_manualtime = STATS_manualtime + 10
				
		'reading
		EMReadScreen diet_mfip1_code, 2, 8, 40
		EMReadScreen diet_mfip1_ver, 1, 8, 51
		EMReadScreen diet_mfip2_code, 2, 9, 40
		EMReadScreen diet_mfip2_ver, 1, 9, 51
		EMReadScreen diet_msa1_code, 2, 11, 40
		EMReadScreen diet_msa1_ver, 1, 11, 51
		EMReadScreen diet_msa2_code, 2, 12, 40
		EMReadScreen diet_msa2_ver, 1, 12, 51
		EMReadScreen diet_msa3_code, 2, 13, 40
		EMReadScreen diet_msa3_ver, 1, 13, 51
		EMReadScreen diet_msa4_code, 2, 14, 40
		EMReadScreen diet_msa4_ver, 1, 14, 51
			
		'writing
		IF diet_mfip1_code <> "__"     THEN CALL check_for_data_validation(diet_row, excel_col, diet_mfip1_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_mfip1_ver <> "_"    THEN CALL check_for_data_validation(diet_row + 1, excel_col, diet_mfip1_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_mfip2_code <> "__" THEN CALL check_for_data_validation(diet_row + 2, excel_col, diet_mfip2_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_mfip2_ver <> "_"    THEN CALL check_for_data_validation(diet_row + 3, excel_col, diet_mfip2_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa1_code <> "__"   THEN CALL check_for_data_validation(diet_row + 4, excel_col, diet_msa1_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa1_ver <> "__"     THEN CALL check_for_data_validation(diet_row + 5, excel_col, diet_msa1_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa2_code <> "__"   THEN CALL check_for_data_validation(diet_row + 6, excel_col, diet_msa2_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa2_ver <> "__"     THEN CALL check_for_data_validation(diet_row + 7, excel_col, diet_msa2_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa3_code <> "__"   THEN CALL check_for_data_validation(diet_row + 8, excel_col, diet_msa3_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa3_ver <> "__"     THEN CALL check_for_data_validation(diet_row + 9, excel_col, diet_msa3_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa4_code <> "__"  THEN CALL check_for_data_validation(diet_row + 10, excel_col, diet_msa4_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF diet_msa4_ver <> "__"    THEN CALL check_for_data_validation(diet_row + 11, excel_col, diet_msa4_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'DISA
CALL navigate_to_MAXIS_screen("STAT", "DISA")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_disa, 1, 2, 78
	IF num_of_disa <> "0" THEN 
		'adding to STATS_manualtime -------------------
		STATS_manualtime = STATS_manualtime + 20
		
		'reading
		EMReadScreen disa_bgn_dt, 10, 6, 47
		EMReadScreen disa_end_dt, 10, 6, 69
		EMReadScreen disa_cert_bgn, 10, 7, 47
		EMReadScreen disa_cert_end, 10, 7, 69
		EMReadScreen disa_eld_wvr_bgn, 10, 8, 47
		EMReadScreen disa_eld_wvr_end, 10, 8, 69
		EMReadScreen disa_grh_plan_bgn, 10, 9, 47
		EMReadScreen disa_grh_plan_end, 10, 9, 69
		EMReadScreen disa_cash_code, 2, 11, 59
		EMReadScreen disa_cash_ver, 1, 11, 69
		EMReadScreen disa_snap_code, 2, 12, 59
		EMReadScreen disa_snap_ver, 1, 12, 69
		EMReadScreen disa_hc_code, 2, 13, 59
		EMReadScreen disa_hc_ver, 1, 13, 69
		EMReadScreen disa_home_comm_wvr, 1, 14, 59
		EMReadScreen disa_drug_alc_ver, 1, 18, 69
		
		'writing
		IF disa_bgn_dt <> "__ __ ____" 			THEN objExcel.Cells(disa_row, excel_col).Value = replace(disa_bgn_dt, " ", "/")
		IF disa_end_dt <> "__ __ ____" 			THEN objExcel.Cells(disa_row + 1, excel_col).Value = replace(disa_end_dt, " ", "/")
		IF disa_cert_bgn <> "__ __ ____" 		THEN objExcel.Cells(disa_row + 2, excel_col).Value = replace(disa_cert_bgn , " ", "/")
		IF disa_cert_end <> "__ __ ____" 		THEN objExcel.Cells(disa_row + 3, excel_col).Value = replace(disa_cert_end , " ", "/")
		IF disa_eld_wvr_bgn <> "__ __ ____" 	THEN objExcel.Cells(disa_row + 4, excel_col).Value = replace(disa_eld_wvr_bgn , " ", "/")
		IF disa_eld_wvr_end <> "__ __ ____" 	THEN objExcel.Cells(disa_row + 5, excel_col).Value = replace(disa_eld_wvr_end , " ", "/")
		IF disa_grh_plan_bgn <> "__ __ ____" 	THEN objExcel.Cells(disa_row + 6, excel_col).Value = replace(disa_grh_plan_bgn , " ", "/")
		IF disa_grh_plan_end <> "__ __ ____" 	THEN objExcel.Cells(disa_row + 7, excel_col).Value = replace(disa_grh_plan_end , " ", "/")
		IF disa_cash_code <> "__" 				THEN      CALL check_for_data_validation(disa_row + 8, excel_col, disa_cash_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_cash_ver <> "_" 				THEN       CALL check_for_data_validation(disa_row + 9, excel_col, disa_cash_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_snap_code <> "__" 				THEN     CALL check_for_data_validation(disa_row + 10, excel_col, disa_snap_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_snap_ver <> "_" 				THEN      CALL check_for_data_validation(disa_row + 11, excel_col, disa_snap_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_hc_code <> "__" 				THEN       CALL check_for_data_validation(disa_row + 12, excel_col, disa_hc_code, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_hc_ver <> "_" 					THEN        CALL check_for_data_validation(disa_row + 13, excel_col, disa_hc_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_home_comm_wvr <> "_" 			THEN CALL check_for_data_validation(disa_row + 14, excel_col, disa_home_comm_wvr, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF disa_drug_alc_ver <> "_" 			THEN objExcel.Cells(disa_row + 15, excel_col).Value = disa_drug_alc_ver
	END IF
	excel_col = excel_col + 1
NEXT

'DSTT
CALL navigate_to_MAXIS_screen("STAT", "DSTT")
EMReadScreen num_of_dstt, 1, 2, 78
IF num_of_dstt <> "0"THEN 
	'adding STATS_manualtime ------------------------
	STATS_manualtime = STATS_manualtime + 8
	
	'reading
	EMReadScreen dstt_ongoing_income, 1, 6, 69
	EMReadScreen dstt_hh_income_stop, 8, 9, 69
		dstt_hh_income_stop = replace(dstt_hh_income_stop, " ", "/")
	EMReadScreen dstt_income_exp_amt, 8, 12, 71
		dstt_income_exp_amt = trim(dstt_income_exp_amt)
		dstt_income_exp_amt = replace(dstt_income_exp_amt, "_", "")
	
	'writing
	CALL check_for_data_validation(dstt_row, 3, dstt_ongoing_income, objExcel, objWorkbook, objTemplate, objNewSheet)
	objExcel.Cells(dstt_row + 1, 3).Value = dstt_hh_income_stop
	objExcel.Cells(dstt_row + 2, 3).Value = dstt_income_exp_amt	
END IF

'EATS
CALL navigate_to_MAXIS_screen("STAT", "EATS")
EMReadScreen num_of_eats, 1, 2, 78
IF num_of_eats <> "0" THEN 
	'adding STATS_manualtime -------------------------
	STATS_manualtime = STATS_manualtime + 10
	
	'reading
	EMReadScreen eats_all_eat_together, 1, 4, 72
	EMReadScreen eats_app_boarder, 1, 5, 72
	EMReadScreen eats_group1, 30, 13, 39
		eats_group1 = replace(eats_group1, "?", "_")
		eats_group1 = replace(eats_group1, "  ", ",")
		eats_group1 = replace(eats_group1, ",__", "")
	EMReadScreen eats_group2, 30, 14, 39
		eats_group2 = replace(eats_group2, "?", "_")
		eats_group2 = replace(eats_group2, "  ", ",")
		eats_group2 = replace(eats_group2, ",__", "")
	EMReadScreen eats_group3, 30, 15, 39
		eats_group3 = replace(eats_group3, "?", "_")
		eats_group3 = replace(eats_group3, "  ", ",")
		eats_group3 = replace(eats_group3, ",__", "")
	
	'writing
	CALL check_for_data_validation(eats_row, 3, eats_all_eat_together, objExcel, objWorkbook, objTemplate, objNewSheet)
	CALL check_for_data_validation(eats_row + 1, 3, eats_app_boarder, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF eats_group1 <> "__" THEN objExcel.Cells(eats_row + 2, 3).Value = eats_group1
	IF eats_group2 <> "__" THEN objExcel.Cells(eats_row + 3, 3).Value = eats_group2
	IF eats_group3 <> "__" THEN objExcel.Cells(eats_row + 4, 3).Value = eats_group3
END IF

'EMMA
CALL navigate_to_MAXIS_screen("STAT", "EMMA")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_emma, 1, 2, 78
	IF num_of_emma <> "0" THEN 
		'adding STATS_manualtime ---------------------
		STATS_manualtime = STATS_manualtime + 15
		
		'reading
		EMReadScreen emma_med_emer, 2, 6, 46
		EMReadScreen emma_health_conseq, 2, 8, 46
		EMReadScreen emma_verif, 2, 10, 46
		EMReadScreen emma_begin_dt, 8, 12, 46
		EMReadScreen emma_end_dt, 8, 14, 46
	
		'writing
		CALL check_for_data_validation(emma_row, excel_col, emma_med_emer, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(emma_row + 1, excel_col, emma_health_conseq, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(emma_row + 2, excel_col, emma_verif)
		IF emma_begin_dt <> "__ __ __" THEN objExcel.Cells(emma_row + 4, excel_col).Value = emma_begin_dt
		IF emma_end_dt <> "__ __ __" THEN objExcel.Cells(emma_row + 5, excel_col).Value = emma_end_dt
	END IF
	excel_col = excel_col + 1
NEXT

'EMPS
CALL navigate_to_MAXIS_screen("STAT", "EMPS")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_emps, 1, 2, 78
	IF num_of_emps <> "0" THEN 
		'adding STATS_manualtime ---------------------------
		STATS_manualtime = STATS_manualtime + 28
		
		'reading
		EMReadScreen emps_orient_date, 8, 5, 39
		EMReadScreen emps_attend_orient, 1, 5, 65
		EMReadScreen emps_good_cause, 2, 5, 79
		EMReadScreen emps_sanc_bgn_dt, 8, 6, 39
		EMReadScreen emps_sanc_end_dt, 8, 6, 65
		EMReadScreen emps_spec_med_care, 1, 8, 76
		EMReadScreen emps_care_of_fam, 1, 9, 76
		EMReadScreen emps_pers_cris, 1, 10, 76
		EMReadScreen emps_hard_employ, 2, 11, 76
		EMReadScreen emps_ft_care, 1, 12, 76
		EMReadScreen emps_dwp_plan_dt, 8, 17, 40
		
		'writing
		IF emps_orient_date <> "__ __ __" 		THEN objExcel.Cells(emps_row, excel_col).Value = replace(emps_orient_date, " ", "/")
		IF emps_attend_orient <> "_" 			THEN CALL check_for_data_validation(emps_row + 1, excel_col, emps_attend_orient, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_good_cause <> "__" 				THEN CALL check_for_data_validation(emps_row + 2, excel_col, emps_good_cause, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_sanc_bgn_dt <> "__ 01 __" 		THEN CALL check_for_data_validation(emps_row + 3, excel_col, emps_sanc_bgn_dt, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_sanc_end_dt <> "__ 01 __" 		THEN CALL check_for_data_validation(emps_row + 4, excel_col, emps_sanc_end_dt, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_spec_med_care <> "_" 			THEN CALL check_for_data_validation(emps_row + 5, excel_col, emps_spec_med_care, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_care_of_fam <> "_" 				THEN CALL check_for_data_validation(emps_row + 6, excel_col, emps_care_of_fam, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_pers_cris <> "_" 				THEN CALL check_for_data_validation(emps_row + 7, excel_col, emps_pers_cris, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_hard_employ <> "__" 			THEN CALL check_for_data_validation(emps_row + 8, excel_col, emps_hard_employ, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_ft_care <> "_"					THEN CALL check_for_data_validation(emps_row + 9, excel_col, emps_ft_care, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF emps_dwp_plan_dt <> "__ __ __"		THEN objExcel.Cells(emps_row + 10, excel_col).Value = replace(emps_dwp_plan_dt, " ", "/")
	END IF
	excel_col = excel_col + 1
NEXT

'FACI
CALL navigate_to_MAXIS_screen("STAT", "FACI")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_faci, 1, 2, 78
	IF num_of_faci <> "0" THEN 
		'adding STATS_manualtime ------------------------------
		STATS_manualtime = STATS_manualtime + 29
		
		'reading
		EMReadScreen faci_vendor_number, 8, 5, 43
		EMReadScreen faci_vendor_name, 30, 6, 43
			faci_vendor_name = trim(replace(faci_vendor_name, "_", ""))
		EMReadScreen faci_type, 2, 7, 43
		EMReadScreen faci_fs_elig, 1, 8, 43
		EMReadScreen faci_fs_faci_type, 1, 8, 71
		EMReadScreen faci_date_in, 10, 14, 47
		EMReadScreen faci_date_out, 10, 14, 71
		
		'writing
		objExcel.Cells(faci_row, excel_col).Value = faci_vendor_number
		objExcel.Cells(faci_row + 1, excel_col).Value = faci_vendor_name
		IF faci_type <> "__" THEN CALL check_for_data_validation(faci_row + 2, excel_col, faci_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF faci_fs_elig <> "_" THEN CALL check_for_data_validation(faci_row + 3, excel_col, faci_fs_elig, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF faci_fs_faci_type <> "_" THEN CALL check_for_data_validation(faci_row + 4, excel_col, faci_fs_faci_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(faci_row + 5, excel_col).Value = replace(faci_date_in, " ", "/")
		objExcel.Cells(faci_row + 6, excel_col).Value = replace(faci_date_out, " ", "/")
	END IF
	excel_col = excel_col + 1
NEXT

'FMED
CALL navigate_to_MAXIS_screen("STAT", "FMED")
excel_col = 3

REDIM fmed_array(3, 6)
'(i, 0) = type
'(i, 1) = verif
'(i, 2) = ref num 
'(i, 3) = category
'(i, 4) = begin date 
'(i, 5) = end date
'(i, 6) = amount

FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_fmed, 1, 2, 78
	IF num_of_fmed <> "0" THEN 
		'reading
		EMReadScreen fmed_miles, 4, 17, 34
			
		fmed_read_row = 9
		FOR i = 0 TO 3
			EMReadScreen fmed_array(i, 0), 2, fmed_read_row + i, 25
			EMReadScreen fmed_array(i, 1), 2, fmed_read_row + i, 32
			EMReadScreen fmed_array(i, 2), 2, fmed_read_row + i, 38
			EMReadScreen fmed_array(i, 3), 1, fmed_read_row + i, 44
			EMReadScreen fmed_array(i, 4), 5, fmed_read_row + i, 50
			EMReadScreen fmed_array(i, 5), 5, fmed_read_row + i, 60
			EMReadScreen fmed_array(i, 6), 8, fmed_read_row + i, 70
				fmed_array(i, 6) = trim(replace(fmed_array(i, 6), "_", ""))
		NEXT
		
		'writing
		objExcel.Cells(fmed_row, excel_col).Value = trim(replace(fmed_miles, "_", ""))
		fmed_write_row = fmed_row + 1
		FOR a = 0 TO 3
			IF fmed_array(a, 0) <> "__" THEN 
				'adding STATS_manualtime only for those FMED rows that are written --------------------------
				STATS_manualtime = STATS_manualtime + 20
				
				CALL check_for_data_validation(fmed_write_row, excel_col, fmed_array(a, 0))
				fmed_write_row = fmed_write_row + 1
				IF fmed_array(a, 1) <> "__" 		THEN CALL check_for_data_validation(fmed_write_row, excel_col, fmed_array(a, 1), objExcel, objWorkbook, objTemplate, objNewSheet)
				fmed_write_row = fmed_write_row + 1
				IF fmed_array(a, 2) <> "__"			THEN objExcel.Cells(fmed_write_row, excel_col).Value = fmed_array(a, 2)
				fmed_write_row = fmed_write_row + 1
				IF fmed_array(a, 3) <> "_"			THEN CALL check_for_data_validation(fmed_write_row, excel_col, fmed_array(a, 3), objExcel, objWorkbook, objTemplate, objNewSheet)
				fmed_write_row = fmed_write_row + 1
				IF fmed_array(a, 4) <> "__ __"		THEN objExcel.Cells(fmed_write_row, excel_col).Value = replace(fmed_array(a, 4), " ", "/")
				fmed_write_row = fmed_write_row + 1
				IF fmed_array(a, 5) <> "__ __"		THEN objExcel.Cells(fmed_write_row, excel_col).Value = replace(fmed_array(a, 5), " ", "/")
				fmed_write_row = fmed_write_row + 1
				IF fmed_array(a, 6) <> ""			THEN objExcel.Cells(fmed_write_row, excel_col).Value = fmed_array(a, 6)
				fmed_write_row = fmed_write_row + 1
			END IF
		NEXT
			
		'reseting the array
		FOR ib = 0 TO 3
			FOR jk = 0 TO 6
				fmed_array(ib, jk) = ""
			NEXT
		NEXT
	END IF
	excel_col = excel_col + 1
NEXT

'HEST
CALL navigate_to_MAXIS_screen("STAT", "HEST")
EMReadScreen num_of_hest, 1, 2, 78
IF num_of_hest <> "0" THEN 
	'adding STATS_manualtime ---------------------
	STATS_manualtime = STATS_manualtime + 16
	
	'reading
	EMReadScreen hest_fs_choice_date, 8, 7, 40
	EMReadScreen hest_initial_month, 8, 8, 61
		hest_initial_month = trim(hest_initial_month)
		hest_initial_month = replace(hest_initial_month, "_", "")
	EMReadScreen hest_retro_heat_air, 1, 13, 34
	EMReadScreen hest_prosp_heat_air, 1, 13, 60
	EMReadScreen hest_retro_elec, 1, 14, 34
	EMReadScreen hest_prosp_elec, 1, 14, 60
	EMReadScreen hest_retro_phone, 1, 15, 34
	EMReadScreen hest_prosp_phone, 1, 15, 60
	
	'writing
	IF hest_fs_choice_date <> "__ __ __" THEN objExcel.Cells(hest_row, 3).Value = hest_fs_choice_date
	IF hest_initial_month <> "" THEN objExcel.Cells(hest_row + 1, 3).Value = hest_initial_month
	IF hest_retro_heat_air <> "_" THEN CALL check_for_data_validation(hest_row + 2, 3, hest_retro_heat_air, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF hest_prosp_heat_air <> "_" THEN CALL check_for_data_validation(hest_row + 3, 3, hest_prosp_heat_air, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF hest_retro_elec <> "_" THEN CALL check_for_data_validation(hest_row + 4, 3, hest_retro_elec, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF hest_prosp_elec <> "_" THEN CALL check_for_data_validation(hest_row + 5, 3, hest_prosp_elec, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF hest_retro_phone <> "_" THEN CALL check_for_data_validation(hest_row + 6, 3, hest_retro_phone, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF hest_prosp_phone <> "_" THEN CALL check_for_data_validation(hest_row + 7, 3, hest_prosp_phone, objExcel, objWorkbook, objTemplate, objNewSheet)
END IF

'IMIG
CALL navigate_to_MAXIS_screen("STAT", "IMIG")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_imig, 1, 2, 78
	IF num_of_imig <> "0" THEN 
		'adding STATS_manualtime ------------------------
		STATS_manualtime = STATS_manualtime + 33
		
		'reading
		EMReadScreen imig_status, 2, 6, 45
		EMReadScreen imig_entry_date, 10, 7, 45
		EMReadScreen imig_status_date, 10, 7, 71
		EMReadScreen imig_status_ver, 2, 8, 45
		EMReadScreen imig_status_lpr, 2, 9, 45
		EMReadScreen imig_nationality, 2, 10, 45
		EMReadScreen imig_40_soc_sec_cr, 1, 13, 56
		EMReadScreen imig_40_soc_sec_ver, 1, 13, 71
		EMReadScreen imig_battered, 1, 14, 56
		EMReadScreen imig_battered_ver, 1, 14, 71
		EMReadScreen imig_lil_green_army_man, 1, 15, 56
		EMReadScreen imig_lil_green_army_man_ver, 1, 15, 71
		EMReadScreen imig_hmong_lao, 2, 16, 56
		EMReadScreen imig_esl_coop, 1, 17, 56
		EMReadScreen imig_esl_coop_ver, 1, 17, 71
		EMReadScreen imig_esl_skillz, 1, 18, 56
		
		'writing
		IF imig_status <> "__" 					THEN CALL check_for_data_validation(imig_row, excel_col, imig_status, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_entry_date <> "__ __ ____" 		THEN objExcel.Cells(imig_row + 1, excel_col).Value = replace(imig_entry_date, " ", "/")
		IF imig_status_date <> "__ __ ____" 	THEN objExcel.Cells(imig_row + 2, excel_col).Value = replace(imig_status_date, " ", "/")
		IF imig_status_ver <> "__" 				THEN CALL check_for_data_validation(imig_row + 3, excel_col, imig_status_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_status_lpr <> "__" 				THEN CALL check_for_data_validation(imig_row + 4, excel_col, imig_status_lpr, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_nationality <> "__" 			THEN CALL check_for_data_validation(imig_row + 5, excel_col, imig_nationality, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_40_soc_sec_cr <> "_" 			THEN CALL check_for_data_validation(imig_row + 6, excel_col, imig_40_soc_sec_cr, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_40_soc_sec_ver <> "_" 			THEN CALL check_for_data_validation(imig_row + 7, excel_col, imig_40_soc_sec_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_battered <> "_" 				THEN CALL check_for_data_validation(imig_row + 8, excel_col, imig_battered, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_battered_ver <> "_" 			THEN CALL check_for_data_validation(imig_row + 9, excel_col, imig_battered_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_lil_green_army_man <> "_" 		THEN CALL check_for_data_validation(imig_row + 10, excel_col, imig_lil_green_army_man, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_lil_green_army_man_ver <> "_" 	THEN CALL check_for_data_validation(imig_row + 11, excel_col, imig_lil_green_army_man_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_hmong_lao <> "__" 				THEN CALL check_for_data_validation(imig_row + 12, excel_col, imig_hmong_lao, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_esl_coop <> "_" 				THEN CALL check_for_data_validation(imig_row + 13, excel_col, imig_esl_coop, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_esl_coop_ver <> "_" 			THEN CALL check_for_data_validation(imig_row + 14, excel_col, imig_esl_coop_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF imig_esl_skillz <> "_" 				THEN CALL check_for_data_validation(imig_row + 15, excel_col, imig_esl_skillz, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'INSA
CALL navigate_to_MAXIS_screen("STAT", "INSA")
excel_col = 3
CALL write_value_and_transmit("01", 20, 79)
EMReadScreen num_of_insa, 1, 2, 78
IF num_of_insa <> "0" THEN 
	'adding STATS_manualtime ----------------------
	STATS_manualtime = STATS_manualtime + 12
	
	'reading
	EMReadScreen insa_resp_coop, 1, 4, 62
	EMReadScreen insa_good_cause_status, 1, 5, 62
	EMReadScreen insa_good_cause_date, 8, 6, 62
	EMReadScreen insa_good_cause_evid, 1, 7, 62
	EMReadScreen insa_good_cause_rqmt, 1, 8, 62
	EMReadScreen insa_company, 38, 10, 38
		insa_company = trim(replace(insa_company, "_", ""))
	EMReadScreen insa_drug_cov, 1, 11, 62
	EMReadScreen insa_cov_end_date, 8, 12, 62
	EMReadScreen insa_covered_dudes1, 38, 15, 30
		insa_covered_dudes1 = replace(insa_covered_dudes1, "  ", ",")
		insa_covered_dudes1 = replace(insa_covered_dudes1, ",__", "")
	EMReadScreen insa_covered_dudes2, 38, 16, 30
		insa_covered_dudes2 = replace(insa_covered_dudes2, "  ", ",")
		insa_covered_dudes2 = replace(insa_covered_dudes2, ",__", "")
	insa_covered_dudes = ""
	IF insa_covered_dudes1 <> "__" THEN insa_covered_dudes = insa_covered_dudes1
	IF insa_covered_dudes2 <> "__" THEN insa_covered_dudes = insa_covered_dudes & "," & insa_covered_dudes2
	
	'writing
	IF insa_resp_coop <> "_" 				THEN CALL check_for_data_validation(insa_row, excel_col, insa_resp_coop, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF insa_good_cause_status <> "_" 		THEN CALL check_for_data_validation(insa_row + 1, excel_col, insa_good_cause_status, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF insa_good_cause_date <> "__ __ __" 	THEN objExcel.Cells(insa_row + 2, excel_col).Value = replace(insa_good_cause_date, " ", "/")
	IF insa_good_cause_evid <> "_"			THEN CALL check_for_data_validation(insa_row + 3, excel_col, insa_good_cause_evid, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF insa_good_cause_rqmt <> "_"			THEN CALL check_for_data_validation(insa_row + 4, excel_col, insa_good_cause_rqmt, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF insa_company <> ""					THEN objExcel.Cells(insa_row + 5, excel_col).Value = insa_company
	IF insa_drug_cov <> "_"					THEN CALL check_for_data_validation(insa_row + 6, excel_col, insa_drug_cov, objExcel, objWorkbook, objTemplate, objNewSheet)
	IF insa_cov_end_date <> "__ __ __"		THEN objExcel.Cells(insa_row + 7, excel_col).Value = replace(insa_cov_end_date, " ", "/")
	IF insa_covered_dudes <> "" 			THEN objExcel.Cells(insa_row + 8, excel_col).Value = insa_covered_dudes
END IF

'JOBS1, JOBS2, JOBS3
CALL navigate_to_MAXIS_screen("STAT", "JOBS")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_jobs, 1, 2, 78
	IF num_of_jobs <> "0" THEN 
		num_of_active_jobs = 0
		DO
			EMReadScreen jobs_end_date, 8, 9, 49
			IF jobs_end_date = "__ __ __" THEN num_of_active_jobs = num_of_active_jobs + 1
			transmit
			EMReadScreen enter_a_valid, 13, 24, 2
		LOOP UNTIL enter_a_valid = "ENTER A VALID"
		
		IF num_of_active_jobs >= 1 THEN
			CALL write_value_and_transmit("01", 20, 79)
			DO
				EMReadScreen jobs_end_date, 8, 9, 49
				IF jobs_end_date = "__ __ __" THEN EXIT DO
				transmit
			LOOP
			EMReadScreen jobs1_type, 1, 5, 38
			EMReadScreen jobs1_verif, 1, 6, 38
			EMReadScreen jobs1_employer, 30, 7, 42
				jobs1_employer = trim(replace(jobs1_employer, "_", ""))
			EMReadScreen jobs1_inc_start, 8, 9, 35
				jobs1_inc_start = replace(jobs1_inc_start, " ", "/")
				
			'if this is a snap case, we will grab earned income information off the PIC
			IF snap_case = TRUE THEN 
				CALL write_value_and_transmit("X", 19, 38)
				EMReadScreen jobs1_pay_freq, 1, 5, 64
				EMReadScreen jobs1_per_pay_period_hrs, 6, 16, 51
					jobs1_per_pay_period_hrs = (trim(jobs1_per_pay_period_hrs) * 1)
				EMReadScreen jobs1_per_pay_period_earn, 8, 17, 56
					jobs1_per_pay_period_earn = (trim(jobs1_per_pay_period_earn) * 1)
				jobs1_income = jobs1_per_pay_period_earn / jobs1_per_pay_period_hrs
				
				IF jobs1_pay_freq = "1" THEN 
					jobs1_weekly_hrs = jobs1_per_pay_period_hrs / 4.3
				ELSEIF jobs1_pay_freq = "2" THEN 
					jobs1_weekly_hrs = jobs1_per_pay_period_hrs / 2.15
				ELSEIF jobs1_pay_freq = "3" THEN 
					jobs1_weekly_hrs = jobs1_per_pay_period_hrs / 2
				ELSEIF jobs1_pay_freq = "4" THEN 
					jobs1_weekly_hrs = jobs1_per_pay_period_hrs
				END IF
				PF3
			END IF
			'if this is a cash case with no snap, the script will grab hourly and weekly income information from the prospective side
			IF snap_case = FALSE THEN 
				EMReadScreen jobs1_pay_freq, 1, 18, 35
				'Reading and converting the total number of hours worked in the prospective budget cycle
				EMReadScreen jobs1_month_hrs, 3, 18, 72
					IF InStr(jobs1_month_hrs, "?") <> 0 THEN 
						jobs1_month_hrs = 1
						MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "The script has found a question mark in this JOBS panel. It will substitue a 1 for working hours but you should update working hours for JOBS1 with an appropriate value AFTER the script has finished running.", vbInformation + vbSystemModal, "Question Mark Found"
					ELSE
						jobs1_month_hrs = trim(jobs1_month_hrs) * 1
						jobs1_weekly_hrs = jobs1_month_hrs / 4.33
					END IF
				'Reading and converting the total income in the prospective budget cycle
				EMReadScreen jobs1_income, 8, 17, 67
					jobs1_income = (trim(jobs1_income) * 1)
					jobs1_income = jobs1_income / jobs1_month_hrs
			END IF
			
			'writing
			CALL check_for_data_validation(jobs1_row, excel_col, jobs1_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(jobs1_row + 1, excel_col, jobs1_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(jobs1_row + 2, excel_col).Value = jobs1_employer
			IF jobs1_inc_start <> "__/__/__" THEN objExcel.Cells(jobs1_row + 3, excel_col).Value = jobs1_inc_start
			objExcel.Cells(jobs1_row + 4, excel_col).Value = jobs1_pay_freq
			objExcel.Cells(jobs1_row + 5, excel_col).Value = FormatNumber(jobs1_weekly_hrs, 2)
			objExcel.Cells(jobs1_row + 6, excel_col).Value = FormatNumber(jobs1_income, 2)
		END IF
		IF num_of_active_jobs >= 2 THEN
			transmit
			DO
				EMReadScreen jobs_end_date, 8, 9, 49
				IF jobs_end_date = "__ __ __" THEN EXIT DO
				transmit
			LOOP
			EMReadScreen jobs2_type, 1, 5, 38
			EMReadScreen jobs2_verif, 1, 6, 38
			EMReadScreen jobs2_employer, 30, 7, 42
				jobs2_employer = trim(replace(jobs2_employer, "_", ""))
			EMReadScreen jobs2_inc_start, 8, 9, 35
				jobs2_inc_start = replace(jobs2_inc_start, " ", "/")
			EMReadScreen jobs2_pay_freq, 1, 18, 35
			'Reading and converting the total number of hours worked in the prospective budget cycle
			EMReadScreen jobs2_month_hrs, 3, 18, 72
				IF InStr(jobs2_month_hrs, "?") <> 0 THEN 
					jobs2_month_hrs = 1
					MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "The script has found a question mark in this JOBS panel. It will substitue a 1 for working hours but you should update working hours for JOBS2 with an appropriate value AFTER the script has finished running.", vbInformation + vbSystemModal, "Question Mark Found"					
				ELSE
					jobs2_month_hrs = trim(jobs2_month_hrs) * 1
					jobs2_weekly_hrs = jobs2_month_hrs / 4.33
				END IF
			'Reading and converting the total income in the prospective budget cycle
			EMReadScreen jobs2_income, 8, 17, 67
				jobs2_income = (trim(jobs2_income) * 1)
				jobs2_income = jobs2_income / jobs2_month_hrs
			
			'writing
			CALL check_for_data_validation(jobs2_row, excel_col, jobs2_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(jobs2_row + 1, excel_col, jobs2_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(jobs2_row + 2, excel_col).Value = jobs2_employer
			objExcel.Cells(jobs2_row + 3, excel_col).Value = jobs2_inc_start
			objExcel.Cells(jobs2_row + 4, excel_col).Value = jobs2_pay_freq
			objExcel.Cells(jobs2_row + 5, excel_col).Value = FormatNumber(jobs2_weekly_hrs, 2)
			objExcel.Cells(jobs2_row + 6, excel_col).Value = FormatNumber(jobs2_income, 2)
		END IF
		IF num_of_active_jobs >= 3 THEN
			transmit
			DO
				EMReadScreen jobs_end_date, 8, 9, 49
				IF jobs_end_date = "__ __ __" THEN EXIT DO
				transmit
			LOOP
			EMReadScreen jobs3_type, 1, 5, 38
			EMReadScreen jobs3_verif, 1, 6, 38
			EMReadScreen jobs3_employer, 30, 7, 42
				jobs3_employer = trim(replace(jobs3_employer, "_", ""))
			EMReadScreen jobs3_inc_start, 8, 9, 35
				jobs3_inc_start = replace(jobs3_inc_start, " ", "/")
			EMReadScreen jobs3_pay_freq, 1, 18, 35
			'Reading and converting the total number of hours worked in the prospective budget cycle
			EMReadScreen jobs3_month_hrs, 3, 18, 72
				IF InStr(jobs3_month_hrs, "?") <> 0 THEN 
					jobs3_month_hrs = 1
					MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & "The script has found a question mark in this JOBS panel. It will substitue a 1 for working hours but you should update working hours for JOBS3 with an appropriate value AFTER the script has finished running.", vbInformation + vbSystemModal, "Question Mark Found"
				ELSE
					jobs3_month_hrs = trim(jobs3_month_hrs) * 1
					jobs3_weekly_hrs = jobs3_month_hrs / 4.33
				END IF
			'Reading and converting the total income in the prospective budget cycle
			EMReadScreen jobs3_income, 8, 17, 67
				jobs3_income = (trim(jobs3_income) * 1)
				jobs3_income = jobs3_income / jobs3_month_hrs
			
			'writing
			CALL check_for_data_validation(jobs3_row, excel_col, jobs3_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(jobs3_row + 1, excel_col, jobs3_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(jobs3_row + 2, excel_col).Value = jobs3_employer
			objExcel.Cells(jobs3_row + 3, excel_col).Value = jobs3_inc_start
			objExcel.Cells(jobs3_row + 4, excel_col).Value = jobs3_pay_freq
			objExcel.Cells(jobs3_row + 5, excel_col).Value = FormatNumber(jobs3_weekly_hrs, 2)
			objExcel.Cells(jobs3_row + 6, excel_col).Value = FormatNumber(jobs3_income, 2)
		END IF		
	END IF
	excel_col = excel_col + 1
NEXT

'MEDI
CALL navigate_to_MAXIS_screen("STAT", "MEDI")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_medi, 1, 2, 78
	IF num_of_medi <> "0" THEN 
		'reading
		EMReadScreen medi_claim_suffix, 3, 6, 56
		EMReadScreen medi_part_a_prem, 8, 7, 46
			medi_part_a_prem = trim(medi_part_a_prem)
			medi_part_a_prem = replace(medi_part_a_prem, "_", "")
		EMReadScreen medi_part_b_prem, 8, 7, 73
			medi_part_b_prem = trim(medi_part_b_prem)
			medi_part_b_prem = replace(medi_part_b_prem, "_", "")
		EMReadScreen medi_part_a_bgn, 8, 15, 24
		EMReadScreen medi_part_b_bgn, 8, 15, 54
		EMReadScreen medi_apply_to_spdn, 1, 11, 71
		EMReadScreen medi_apply_thru_dt, 5, 12, 71
		
		'writing
		objExcel.Cells(medi_row, excel_col).Value = medi_claim_suffix
		objExcel.Cells(medi_row + 1, excel_col).Value = medi_part_a_prem
		objExcel.Cells(medi_row + 2, excel_col).Value = medi_part_b_prem
		IF medi_part_a_bgn <> "__ __ __" THEN objExcel.Cells(medi_row + 3, excel_col).Value = replace(medi_part_a_bgn, " ", "/")
		IF medi_part_b_bgn <> "__ __ __" THEN objExcel.Cells(medi_row + 4, excel_col).Value = replace(medi_part_b_bgn, " ", "/")
		CALL check_for_data_validation(medi_row + 5, excel_col, medi_apply_to_spdn, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(medi_row + 6, excel_col).Value = medi_apply_thru_dt		
	END IF
	excel_col = excel_col + 1
NEXT

'MMSA
CALL navigate_to_MAXIS_screen("STAT", "MMSA")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_mmsa, 1, 2, 78
	IF num_of_mmsa <> "0" THEN 
		'reading
		EMReadScreen mmsa_fed_liv_arr, 1, 7, 54
		EMReadScreen mmsa_cont_elig, 1, 9, 54
		EMReadScreen mmsa_spousal_income, 1, 12, 62
		EMReadScreen mmsa_shared_housing, 1, 14, 62
		
		'writing
		IF mmsa_fed_liv_arr <> "_" 		THEN CALL check_for_data_validation(mmsa_row, excel_col, mmsa_fed_liv_arr, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF mmsa_cont_elig <> "_" 		THEN CALL check_for_data_validation(mmsa_row + 1, excel_col, mmsa_cont_elig, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF mmsa_spousal_income <> "_" 	THEN CALL check_for_data_validation(mmsa_row + 2, excel_col, mmsa_spousal_income, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF mmsa_shared_housing <> "_" 	THEN CALL check_for_data_validation(mmsa_row + 3, excel_col, mmsa_shared_housing, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'MSUR
CALL navigate_to_MAXIS_screen("STAT", "MSUR")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_msur, 1, 2, 78
	IF num_of_msur <> "0" THEN
		'reading
		EMReadScreen msur_start_dt, 10, 7, 36
		msur_start_dt = replace(msur_start_dt, "/", "")
		
		'writing
		objExcel.Cells(msur_row, excel_col).Value = replace(msur_start_dt, " ", "/")
	END IF
	excel_col = excel_col + 1
NEXT

'OTHR
CALL navigate_to_MAXIS_screen("STAT", "OTHR")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_othr, 1, 2, 78
	IF num_of_othr <> "0" THEN 
		'reading
		EMReadScreen othr_asset_type, 1, 6, 40
		EMReadScreen othr_cash_value, 10, 8, 40
			othr_cash_value = trim(othr_cash_value)
			othr_cash_value = replace(othr_cash_value, "_", "")
		EMReadScreen othr_cash_val_ver, 1, 8, 57
		EMReadScreen othr_amt_owed, 10, 9, 40
			othr_amt_owed = trim(othr_amt_owed)
			othr_amt_owed = replace(othr_amt_owed, "_", "")
		EMReadScreen othr_amt_owed_ver, 1, 9, 57
		EMReadScreen othr_verif_dt, 8, 10, 39
		EMReadScreen othr_count_cash, 1, 12, 50
		EMReadScreen othr_count_snap, 1, 12, 57
		EMReadScreen othr_count_hc, 1, 12, 64
		EMReadScreen othr_count_ive, 1, 12, 73
		EMReadScreen othr_joint_own, 1, 13, 44
		EMReadScreen othr_share_ratio, 5, 15, 50
		
		'writing
		CALL check_for_data_validation(othr_row, excel_col, othr_asset_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(othr_row + 1, excel_col).Value = othr_cash_value
		IF othr_cash_val_ver <> "_" THEN CALL check_for_data_validation(othr_row + 2, excel_col, othr_cash_val_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(othr_row + 3, excel_col).Value = othr_amt_owed
		IF othr_amt_owed_ver <> "_" THEN CALL check_for_data_validation(othr_row + 4, excel_col, othr_amt_owed_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF othr_verif_dt <> "__ __ __" THEN objExcel.Cells(othr_row + 5, excel_col).Value = replace(othr_verif_dt, " ", "/")
		IF othr_count_cash <> "_" 	THEN CALL check_for_data_validation(othr_row + 6, excel_col, othr_count_cash, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF othr_count_snap <> "_" 	THEN CALL check_for_data_validation(othr_row + 7, excel_col, othr_count_snap, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF othr_count_hc <> "_" 	THEN CALL check_for_data_validation(othr_row + 8, excel_col, othr_count_hc, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF othr_count_ive <> "_" 	THEN CALL check_for_data_validation(othr_row + 9, excel_col, othr_count_ive, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(othr_row + 10, excel_col, othr_joint_own, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(othr_row + 11, excel_col).Value = cstr(othr_share_ratio)
	END IF
	excel_col = excel_col + 1
NEXT

'PARE
CALL navigate_to_MAXIS_screen("STAT", "PARE")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_pare, 1, 2, 78
	IF num_of_pare <> "0" THEN 
		'reading
		EMReadScreen pare_child1_ref, 2, 8, 24
			pare_child1_ref = replace(pare_child1_ref, "?", "_")
		EMReadScreen pare_child1_rel, 1, 8, 53
		EMReadScreen pare_child1_ver, 2, 8, 71
		EMReadScreen pare_child2_ref, 2, 9, 24
			pare_child2_ref = replace(pare_child2_ref, "?", "_")
		EMReadScreen pare_child2_rel, 1, 9, 53
		EMReadScreen pare_child2_ver, 2, 9, 71
		EMReadScreen pare_child3_ref, 2, 10, 24
			pare_child3_ref = replace(pare_child3_ref, "?", "_")
		EMReadScreen pare_child3_rel, 1, 10, 53
		EMReadScreen pare_child3_ver, 2, 10, 71
		EMReadScreen pare_child4_ref, 2, 11, 24
			pare_child4_ref = replace(pare_child4_ref, "?", "_")
		EMReadScreen pare_child4_rel, 1, 11, 53
		EMReadScreen pare_child4_ver, 2, 11, 71
		EMReadScreen pare_child5_ref, 2, 12, 24
			pare_child5_ref = replace(pare_child5_ref, "?", "_")
		EMReadScreen pare_child5_rel, 1, 12, 53
		EMReadScreen pare_child5_ver, 2, 12, 71
		EMReadScreen pare_child6_ref, 2, 13, 24
			pare_child6_ref = replace(pare_child6_ref, "?", "_")
		EMReadScreen pare_child6_rel, 1, 13, 53
		EMReadScreen pare_child6_ver, 2, 13, 71		
		
		'writing
		IF pare_child1_ref <> "__" THEN 
			objExcel.Cells(pare_row, excel_col).Value = pare_child1_ref
			CALL check_for_data_validation(pare_row + 1, excel_col, pare_child1_rel, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(pare_row + 2, excel_col, pare_child1_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		END IF
		IF pare_child2_ref <> "__" THEN 
			objExcel.Cells(pare_row + 3, excel_col).Value = pare_child2_ref
			CALL check_for_data_validation(pare_row + 4, excel_col, pare_child2_rel, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(pare_row + 5, excel_col, pare_child2_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		END IF
		IF pare_child3_ref <> "__" THEN 
			objExcel.Cells(pare_row + 6, excel_col).Value = pare_child3_ref
			CALL check_for_data_validation(pare_row + 7, excel_col, pare_child3_rel, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(pare_row + 8, excel_col, pare_child3_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		END IF
		IF pare_child4_ref <> "__" THEN 
			objExcel.Cells(pare_row + 9, excel_col).Value = pare_child4_ref
			CALL check_for_data_validation(pare_row + 10, excel_col, pare_child4_rel, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(pare_row + 11, excel_col, pare_child4_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		END IF
		IF pare_child5_ref <> "__" THEN 
			objExcel.Cells(pare_row + 12, excel_col).Value = pare_child5_ref
			CALL check_for_data_validation(pare_row + 13, excel_col, pare_child5_rel, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(pare_row + 14, excel_col, pare_child5_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		END IF
		IF pare_child6_ref <> "__" THEN 
			objExcel.Cells(pare_row + 15, excel_col).Value = pare_child6_ref
			CALL check_for_data_validation(pare_row + 16, excel_col, pare_child6_rel, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(pare_row + 17, excel_col, pare_child6_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		END IF		
	END IF
	excel_col = excel_col + 1
NEXT

'PBEN1, PBEN2, PBEN3
CALL navigate_to_MAXIS_screen("STAT", "PBEN")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_pben, 1, 2, 78
	IF num_of_pben <> "0" THEN 
		EMReadScreen pben1_type, 2, 8, 24
		IF pben1_type <> "__" THEN 
			EMReadScreen pben1_referral_dt, 8, 8, 40
			EMReadScreen pben1_appl_dt, 8, 8, 51
			EMReadScreen pben1_appl_verif, 1, 8, 62
			EMReadScreen pben1_iaa_date, 8, 8, 66
			EMReadScreen pben1_disp_code, 1, 8, 77
			
			IF pben1_referral_dt <> "__ __ __" THEN objExcel.Cells(pben1_row, excel_col).Value = replace(pben1_referral_dt, " ", "/")
			CALL check_for_data_validation(pben1_row + 1, excel_col, pben1_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			IF pben1_appl_dt <> "__ __ __" THEN objExcel.Cells(pben1_row + 2, excel_col).Value = replace(pben1_appl_dt, " ", "/")
			CALL check_for_data_validation(pben1_row + 3, excel_col, pben1_appl_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			IF pben1_iaa_date <> "__ __ __" THEN objExcel.Cells(pben1_row + 4, excel_col).Value = replace(pben1_iaa_date, " ", "/")
			CALL check_for_data_validation(pben1_row + 5, excel_col, pben1_disp_code, objExcel, objWorkbook, objTemplate, objNewSheet)		
		END IF
		EMReadScreen pben2_type, 2, 9, 24
		IF pben2_type <> "__" THEN 
			EMReadScreen pben2_referral_dt, 8, 9, 40
			EMReadScreen pben2_appl_dt, 8, 9, 51
			EMReadScreen pben2_appl_verif, 1, 9, 62
			EMReadScreen pben2_iaa_date, 8, 9, 66
			EMReadScreen pben2_disp_code, 1, 9, 77
			
			IF pben2_referral_dt <> "__ __ __" THEN objExcel.Cells(pben2_row, excel_col).Value = replace(pben2_referral_dt, " ", "/")
			CALL check_for_data_validation(pben2_row + 1, excel_col, pben2_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			IF pben2_appl_dt <> "__ __ __" THEN objExcel.Cells(pben2_row + 2, excel_col).Value = replace(pben2_appl_dt, " ", "/")
			CALL check_for_data_validation(pben2_row + 3, excel_col, pben2_appl_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			IF pben2_iaa_date <> "__ __ __" THEN objExcel.Cells(pben2_row + 4, excel_col).Value = replace(pben2_iaa_date, " ", "/")
			CALL check_for_data_validation(pben2_row + 5, excel_col, pben2_disp_code, objExcel, objWorkbook, objTemplate, objNewSheet)		
		END IF
		EMReadScreen pben3_type, 2, 10, 24
		IF pben3_type <> "__" THEN 
			EMReadScreen pben3_referral_dt, 8, 10, 40
			EMReadScreen pben3_appl_dt, 8, 10, 51
			EMReadScreen pben3_appl_verif, 1, 10, 62
			EMReadScreen pben3_iaa_date, 8, 10, 66
			EMReadScreen pben3_disp_code, 1, 10, 77
			
			IF pben3_referral_dt <> "__ __ __" 		THEN objExcel.Cells(pben3_row, excel_col).Value = replace(pben3_referral_dt, " ", "/")
			CALL check_for_data_validation(pben3_row + 1, excel_col, pben3_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			IF pben3_appl_dt <> "__ __ __" 			THEN objExcel.Cells(pben3_row + 2, excel_col).Value = replace(pben3_appl_dt, " ", "/")
			CALL check_for_data_validation(pben3_row + 3, excel_col, pben3_appl_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			IF pben3_iaa_date <> "__ __ __" 		THEN objExcel.Cells(pben3_row + 4, excel_col).Value = replace(pben3_iaa_date, " ", "/")
			CALL check_for_data_validation(pben3_row + 5, excel_col, pben3_disp_code, objExcel, objWorkbook, objTemplate, objNewSheet)		
		END IF		
	END IF
	excel_col = excel_col + 1
NEXT

'PDED
CALL navigate_to_MAXIS_screen("STAT", "PDED")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_pded, 1, 2, 78
	IF num_of_pded <> "0" THEN 
		EMReadScreen pded_disa_widow, 1, 7, 60
		EMReadScreen pded_adult_child, 1, 8, 60
		EMReadScreen pded_widow_dis, 1, 9, 60
		CALL write_value_and_transmit("X", 10, 25)
			EMReadScreen pded_unea_reason, 15, 10, 51
				pded_unea_reason = trim(pded_unea_reason)
				pded_unea_reason = replace(pded_unea_reason, "_", "")
			PF3
		EMReadScreen pded_unea_ded, 8, 10, 62
			pded_unea_ded = trim(pded_unea_ded)
			pded_unea_ded = replace(pded_unea_ded, "_", "")
		CALL write_value_and_transmit("X", 11, 27)
			EMReadScreen pded_earn_inc_reason, 15, 10, 51
				pded_earn_inc_reason = trim(pded_earn_inc_reason)
				pded_earn_inc_reason = replace(pded_earn_inc_reason, "_", "")
			PF3
		EMReadScreen pded_earn_inc_ded, 8, 11, 62
			pded_earn_inc_ded = trim(pded_earn_inc_ded)
			pded_earn_inc_ded = replace(pded_earn_inc_ded, "_", "")
		EMReadScreen pded_maepd_limit, 1, 12, 65
		EMReadScreen pded_guard_fee, 8, 15, 44
			pded_guard_fee = trim(pded_guard_fee)
			pded_guard_fee = replace(pded_guard_fee, "_", "")
		EMReadScreen pded_rep_payee, 8, 15, 70
			pded_rep_payee = trim(pded_rep_payee)
			pded_rep_payee = replace(pded_rep_payee, "_", "")
		EMReadScreen pded_other_exp, 8, 18, 41
			pded_other_exp = trim(pded_other_exp)
			pded_other_exp = replace(pded_other_exp, "_", "")
		EMReadScreen pded_shel_spec_need, 1, 18, 78
		EMReadScreen pded_excess_need, 8, 19, 41
			pded_excess_need = trim(pded_excess_need)
			pded_excess_need = replace(pded_excess_need, "_", "")
		EMReadScreen pded_straunt, 1, 19, 78
		
		IF pded_disa_widow <> "_" 		THEN CALL check_for_data_validation(pded_row, excel_col, pded_disa_widow, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF pded_adult_child <> "_" 		THEN CALL check_for_data_validation(pded_row + 1, excel_col, pded_adult_child, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF pded_widow_dis <> "_"		THEN CALL check_for_data_validation(pded_row + 2, excel_col, pded_widow_dis, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(pded_row + 3, excel_col).Value = pded_unea_reason
		objExcel.Cells(pded_row + 4, excel_col).Value = pded_unea_ded
		objExcel.Cells(pded_row + 5, excel_col).Value = pded_earn_inc_reason
		objExcel.Cells(pded_row + 6, excel_col).Value = pded_earn_inc_ded
		IF pded_maepd_limit <> "_" 		THEN CALL check_for_data_validation(pded_row + 7, excel_col, pded_maepd_limit, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(pded_row + 8, excel_col).Value = pded_guard_fee
		objExcel.Cells(pded_row + 9, excel_col).Value = pded_rep_payee
		objExcel.Cells(pded_row + 10, excel_col).Value = pded_other_exp
		IF pded_shel_spec_need <> "_"	THEN CALL check_for_data_validation(pded_row + 11, excel_col, pded_shel_spec_need, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(pded_row + 12, excel_col).Value = pded_excess_need
		IF pded_straunt <> "_" 			THEN CALL check_for_data_validation(pded_row + 13, excel_col, pded_straunt, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'PREG
CALL navigate_to_MAXIS_screen("STAT", "PREG")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_preg, 1, 2, 78
	IF num_of_preg <> "0" THEN 
		EMReadScreen conception_dt, 8, 6, 53
			conception_dt = replace(conception_dt, " ", "/")
		EMReadScreen conception_dt_verif, 1, 6, 75
		EMReadScreen third_tri_verif, 1, 8, 75
		EMReadScreen baby_due_date, 8, 10, 53
			baby_due_date = replace(baby_due_date, " ", "/")
		EMReadScreen multiple_birth, 1, 12, 53
		
		objExcel.Cells(preg_row, excel_col).Value = conception_dt
		objExcel.Cells(preg_row + 1, excel_col).Value = conception_dt_verif
		objExcel.Cells(preg_row + 2, excel_col).Value = third_tri_verif
		objExcel.Cells(preg_row + 3, excel_col).Value = baby_due_date
		objExcel.Cells(preg_row + 4, excel_col).Value = multiple_birth		
	END IF
	excel_col = excel_col + 1
NEXT

'RBIC
CALL navigate_to_MAXIS_screen("STAT", "RBIC")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_rbic, 1, 2, 78
	IF num_of_rbic <> "0" THEN
		DO
			'finding the first active RBIC panel
			EMReadScreen panel_end_date, 8, 6, 68
			IF panel_end_date = "__ __ __" THEN 
				'reading
				EMReadScreen rbic_income_type, 2, 5, 44
				EMReadScreen rbic_start_date, 8, 6, 44
				'skipping reading the RBIC end date
				EMReadScreen rbic_group1, 17, 10, 25
					rbic_group1 = replace(rbic_group1, " ", ",")
					rbic_group1 = replace(rbic_group1, ",__", "")
				EMReadScreen rbic_group1_retro, 8, 10, 47
					rbic_group1_retro = trim(replace(rbic_group1_retro, "_", ""))
				EMReadScreen rbic_group1_prosp, 8, 10, 62
					rbic_group1_prosp = trim(replace(rbic_group1_prosp, "_", ""))
				EMReadScreen rbic_group1_verif, 1, 10, 76
				
				EMReadScreen rbic_group2, 17, 11, 25
					rbic_group2 = replace(rbic_group2, " ", ",")
					rbic_group2 = replace(rbic_group2, ",__", "")
				EMReadScreen rbic_group2_retro, 8, 11, 47
					rbic_group2_retro = trim(replace(rbic_group2_retro, "_", ""))
				EMReadScreen rbic_group2_prosp, 8, 11, 62
					rbic_group2_prosp = trim(replace(rbic_group2_prosp, "_", ""))
				EMReadScreen rbic_group2_verif, 1, 11, 76
					
				EMReadScreen rbic_group3, 17, 12, 25
					rbic_group3 = replace(rbic_group3, " ", ",")
					rbic_group3 = replace(rbic_group3, ",__", "")
				EMReadScreen rbic_group3_retro, 8, 12, 47
					rbic_group3_retro = trim(replace(rbic_group3_retro, "_", ""))
				EMReadScreen rbic_group3_prosp, 8, 12, 62
					rbic_group3_prosp = trim(replace(rbic_group3_prosp, "_", ""))
				EMReadScreen rbic_group3_verif, 1, 12, 76
				
				EMReadScreen rbic_retro_hrs, 3, 13, 15
					rbic_retro_hrs = trim(replace(rbic_retro_hrs, "_", ""))
				EMReadScreen rbic_prosp_hrs, 3, 13, 67
					rbic_prosp_hrs = trim(replace(rbic_prosp_hrs, "_", ""))			
				
				EMReadScreen rbic_expense1_type, 2, 15, 25
				EMReadScreen rbic_expense1_retro, 8, 15, 47
					rbic_expense1_retro = trim(replace(rbic_expense1_retro, "_", ""))
				EMReadScreen rbic_expense1_prosp, 8, 15, 62
					rbic_expense1_prosp = trim(replace(rbic_expense1_prosp, "_", ""))
				EMReadScreen rbic_expense1_verif, 1, 15, 76

				EMReadScreen rbic_expense2_type, 2, 16, 25
				EMReadScreen rbic_expense2_retro, 8, 16, 47
					rbic_expense2_retro = trim(replace(rbic_expense2_retro, "_", ""))
				EMReadScreen rbic_expense2_prosp, 8, 16, 62
					rbic_expense2_prosp = trim(replace(rbic_expense2_prosp, "_", ""))
				EMReadScreen rbic_expense2_verif, 1, 16, 76				
				
				'writing
				IF rbic_income_type <> "__" 			THEN CALL check_for_data_validation(rbic_row, excel_col, rbic_income_type, objExcel, objWorkbook, objTemplate, objNewSheet)
				IF rbic_start_date <> "__ __ __" 		THEN objExcel.Cells(rbic_row + 1, excel_col).Value = replace(rbic_start_date, " ", "/")
				'not writing the rbic end date... 
				IF rbic_group1 <> "__" 					THEN objExcel.Cells(rbic_row + 3, excel_col).Value = rbic_group1
				IF rbic_group1_retro <> "" 				THEN objExcel.Cells(rbic_row + 4, excel_col).Value = rbic_group1_retro
				IF rbic_group1_prosp <> "" 				THEN objExcel.Cells(rbic_row + 5, excel_col).Value = rbic_group1_prosp
				IF rbic_group1_verif <> "_"				THEN CALL check_for_data_validation(rbic_row + 6, excel_col, rbic_group1_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				IF rbic_group2 <> "__" 					THEN objExcel.Cells(rbic_row + 7, excel_col).Value = rbic_group2
				IF rbic_group2_retro <> "" 				THEN objExcel.Cells(rbic_row + 8, excel_col).Value = rbic_group2_retro
				IF rbic_group2_prosp <> "" 				THEN objExcel.Cells(rbic_row + 9, excel_col).Value = rbic_group2_prosp
				IF rbic_group2_verif <> "_"				THEN CALL check_for_data_validation(rbic_row + 10, excel_col, rbic_group2_verif, objExcel, objWorkbook, objTemplate, objNewSheet)				
				IF rbic_group3 <> "__" 					THEN objExcel.Cells(rbic_row + 11, excel_col).Value = rbic_group3
				IF rbic_group3_retro <> "" 				THEN objExcel.Cells(rbic_row + 12, excel_col).Value = rbic_group3_retro
				IF rbic_group3_prosp <> "" 				THEN objExcel.Cells(rbic_row + 13, excel_col).Value = rbic_group3_prosp
				IF rbic_group3_verif <> "_"				THEN CALL check_for_data_validation(rbic_row + 14, excel_col, rbic_group3_verif, objExcel, objWorkbook, objTemplate, objNewSheet)				
				IF rbic_retro_hrs <> ""  				THEN objExcel.Cells(rbic_row + 15, excel_col).Value = rbic_retro_hrs
				IF rbic_prosp_hrs <> ""  				THEN objExcel.Cells(rbic_row + 16, excel_col).Value = rbic_prosp_hrs
				IF rbic_expense1_type <> "__"			THEN CALL check_for_data_validation(rbic_row + 17, excel_col, rbic_expense1_type, objExcel, objWorkbook, objTemplate, objNewSheet)
				IF rbic_expense1_retro <> ""			THEN objExcel.Cells(rbic_row + 18, excel_col).Value = rbic_expense1_retro
				IF rbic_expense1_prosp <> ""			THEN objExcel.Cells(rbic_row + 19, excel_col).Value = rbic_expense1_prosp
				IF rbic_expense1_verif <> "_"			THEN CALL check_for_data_validation(rbic_row + 20, excel_col, rbic_expense1_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				IF rbic_expense2_type <> "__"			THEN CALL check_for_data_validation(rbic_row + 21, excel_col, rbic_expense2_type, objExcel, objWorkbook, objTemplate, objNewSheet)
				IF rbic_expense2_retro <> ""			THEN objExcel.Cells(rbic_row + 22, excel_col).Value = rbic_expense2_retro
				IF rbic_expense2_prosp <> ""			THEN objExcel.Cells(rbic_row + 23, excel_col).Value = rbic_expense2_prosp
				IF rbic_expense2_verif <> "_"			THEN CALL check_for_data_validation(rbic_row + 24, excel_col, rbic_expense2_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				
				'all done
				EXIT DO
			ELSE
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
				IF enter_a_valid = "ENTER A VALID" THEN EXIT DO
			END IF
		LOOP
	END IF
	excel_col = excel_col + 1
NEXT

'REST
CALL navigate_to_MAXIS_screen("STAT", "REST")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_rest, 1, 2, 78
	IF num_of_rest <> "0" THEN 
		'reading
		EMReadScreen rest_ppty_type, 1, 6, 39
		EMReadScreen rest_type_verif, 2, 6, 62
		EMReadScreen rest_mark_val, 10, 8, 41
			rest_mark_val = trim(replace(rest_mark_val, "_", ""))
		EMReadScreen rest_value_ver, 2, 8, 62
		EMReadScreen rest_amt_owed, 10, 9, 41
			rest_amt_owed = trim(replace(rest_amt_owed, "_", ""))
		EMReadScreen rest_amt_owed_ver, 2, 9, 62
		EMReadScreen rest_as_of_date, 8, 10, 39
		EMReadScreen rest_ppty_status, 1, 12, 54
		EMReadScreen rest_joint_own, 1, 13, 54
		EMReadScreen rest_share_ratio, 5, 14, 54
		EMReadScreen rest_repay_date, 8, 16, 62
				
		'writing
		IF rest_ppty_type <> "_" 			THEN CALL check_for_data_validation(rest_row, excel_col, rest_ppty_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF rest_type_verif <> "__" 			THEN CALL check_for_data_validation(rest_row + 1, excel_col, rest_type_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(rest_row + 2, excel_col).Value = rest_mark_val
		IF rest_value_ver <> "__" 			THEN CALL check_for_data_validation(rest_row + 3, excel_col, rest_value_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(rest_row + 4, excel_col).Value = rest_amt_owed
		IF rest_amt_owed_ver <> "__" 		THEN CALL check_for_data_validation(rest_row + 5, excel_col, rest_amt_owed_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF rest_as_of_date <> "__ __ __" 	THEN objExcel.Cells(rest_row + 6, excel_col).Value = replace(rest_as_of_date, " ", "/")
		IF rest_ppty_status <> "_" 			THEN CALL check_for_data_validation(rest_row + 7, excel_col, rest_ppty_status, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF rest_joint_own <> "_" 			THEN CALL check_for_data_validation(rest_row + 8, excel_col, rest_joint_own, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(rest_row + 9, excel_col).Value = rest_share_ratio
		IF rest_repay_date <> "__ __ __" 	THEN objExcel.Cells(rest_row + 10, excel_col).Value = replace(rest_repay_date, " ", "/")
	END IF
	excel_col = excel_col + 1
NEXT

'SCHL
CALL navigate_to_MAXIS_screen("STAT", "SCHL")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_schl, 1, 2, 78
	IF num_of_schl <> "0" THEN 
		EMReadScreen schl_status, 1, 6, 40
		EMReadScreen schl_verif, 2, 6, 63
		EMReadScreen schl_type, 2, 7, 40
		EMReadScreen schl_district, 4, 8, 40
		EMReadScreen schl_k_start, 8, 10, 63
		EMReadScreen schl_grad_dt, 5, 11, 63
		EMReadScreen schl_grad_ver, 2, 12, 63
		EMReadScreen schl_funding, 1, 14, 63
		EMReadScreen schl_fs_elig, 2, 16, 63
		EMReadScreen schl_higher_ed, 1, 18, 63
		
		CALL check_for_data_validation(schl_row, excel_col, schl_status, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF schl_verif <> "__" 			THEN CALL check_for_data_validation(schl_row + 1, excel_col, schl_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF schl_type <> "__" 			THEN CALL check_for_data_validation(schl_row + 2, excel_col, schl_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF schl_district <> "____" 		THEN objExcel.Cells(schl_row + 3, excel_col).Value = schl_district
		IF schl_k_start <> "__ __ __" 	THEN objExcel.Cells(schl_row + 4, excel_col).Value = replace(schl_k_start, " ", "/")
		IF schl_grad_dt <> "__ __" 		THEN objExcel.Cells(schl_row + 5, excel_col).Value = replace(schl_grad_dt, " ", "/")
		IF schl_grad_ver <> "__" 		THEN CALL check_for_data_validation(schl_row + 6, excel_col, schl_grad_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF schl_funding <> "_" 			THEN CALL check_for_data_validation(schl_row + 7, excel_col, schl_funding, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF schl_fs_elig <> "__" 		THEN CALL check_for_data_validation(schl_row + 8, excel_col, schl_fs_elig, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF schl_higher_ed <> "_" 		THEN CALL check_for_data_validation(schl_row + 9, excel_col, schl_higher_ed, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'SECU
CALL navigate_to_MAXIS_screen("STAT", "SECU")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_SECU, 1, 2, 78
	IF num_of_SECU <> "0" THEN 
		EMReadScreen secu_type, 2, 6, 50
		EMReadScreen secu_policy, 12, 7, 50
			secu_policy = trim(secu_policy)
			secu_policy = replace(secu_policy, "_", "")
		EMReadScreen secu_name, 20, 8, 50
			secu_name = trim(replace(secu_name, "_", ""))
		EMReadScreen secu_csv, 8, 10, 52
			secu_csv = trim(replace(secu_csv, "_", ""))
		EMReadScreen secu_as_of_date, 8, 11, 35
			secu_as_of_date = replace(secu_as_of_date, " ", "/")
		EMReadScreen secu_as_of_verif, 1, 11, 50
		EMReadScreen secu_face_value, 8, 12, 52
			secu_face_value = trim(replace(secu_face_value, "_", ""))
		EMReadScreen secu_withdraw_pen, 8, 13, 52
			secu_withdraw_pen = trim(replace(secu_withdraw_pen, "_", ""))
		EMReadScreen secu_count_cash, 1, 15, 50
		EMReadScreen secu_count_snap, 1, 15, 57
		EMReadScreen secu_count_hc, 1, 15, 64
		EMReadScreen secu_count_grh, 1, 15, 72
		EMReadScreen secu_count_ive, 1, 15, 80
		EMReadScreen secu_joint_own, 1, 16, 44
		EMReadScreen secu_own_ratio, 5, 16, 76
			
		CALL check_for_data_validation(secu_row, excel_col, secu_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(secu_row + 1, excel_col).Value = secu_policy
		objExcel.Cells(secu_row + 2, excel_col).Value = secu_name
		objExcel.Cells(secu_row + 3, excel_col).Value = secu_csv
		IF secu_as_of_date <> "__/__/__" 	THEN objExcel.Cells(secu_row + 4, excel_col).Value = secu_as_of_date
		IF secu_as_of_verif <> "_" 			THEN CALL check_for_data_validation(secu_row + 5, excel_col, secu_as_of_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(secu_row + 6, excel_col).Value = secu_face_value
		objExcel.Cells(secu_row + 7, excel_col).Value = secu_withdraw_pen
		IF secu_count_cash <> "_" 			THEN CALL check_for_data_validation(secu_row + 8, excel_col, secu_count_cash, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF secu_count_snap <> "_" 			THEN CALL check_for_data_validation(secu_row + 9, excel_col, secu_count_snap, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF secu_count_hc <> "_" 			THEN CALL check_for_data_validation(secu_row + 10, excel_col, secu_count_hc, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF secu_count_grh <> "_" 			THEN CALL check_for_data_validation(secu_row + 11, excel_col, secu_count_grh, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF secu_count_ive <> "_" 			THEN CALL check_for_data_validation(secu_row + 12, excel_col, secu_count_ive, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF secu_joint_own <> "_"			THEN CALL check_for_data_validation(secu_row + 13, excel_col, secu_joint_own, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(secu_row + 14, excel_col).Value = secu_own_ratio		
	END IF
	excel_col = excel_col + 1
NEXT

'SHEL
CALL navigate_to_MAXIS_screen("STAT", "SHEL")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_shel, 1, 2, 78
	IF num_of_shel <> "0" THEN 
		EMReadScreen shel_subsidized, 1, 6, 46
		EMReadScreen shel_shared, 1, 6, 64
		EMReadScreen shel_paid_to, 25, 7, 50
			shel_paid_to = trim(replace(shel_paid_to, "_", ""))
		EMReadScreen shel_retro_rent, 8, 11, 37
			shel_retro_rent = trim(replace(shel_retro_rent, "_", ""))
		EMReadScreen shel_retro_rent_ver, 2, 11, 48
		EMReadScreen shel_prosp_rent, 8, 11, 56
			shel_prosp_rent = trim(replace(shel_prosp_rent, "_", ""))
		EMReadScreen shel_prosp_rent_ver, 2, 11, 67
		EMReadScreen shel_retro_lot_rent, 8, 12, 37
			shel_retro_lot_rent = trim(replace(shel_retro_lot_rent, "_", ""))
		EMReadScreen shel_retro_lot_rent_ver, 2, 12, 48
		EMReadScreen shel_prosp_lot_rent, 8, 12, 56
			shel_prosp_lot_rent = trim(replace(shel_prosp_lot_rent, "_", ""))
		EMReadScreen shel_prosp_lot_rent_ver, 2, 12, 67
		EMReadScreen shel_retro_mortgage, 8, 13, 37
			shel_retro_mortgage = trim(replace(shel_retro_mortgage, "_", ""))
		EMReadScreen shel_retro_mortgage_ver, 2, 13, 48
		EMReadScreen shel_prosp_mortgage, 8, 13, 56
			shel_prosp_mortgage = trim(replace(shel_prosp_mortgage, "_", ""))
		EMReadScreen shel_prosp_mortgage_ver, 2, 13, 67
		EMReadScreen shel_retro_insur, 8, 14, 37
			shel_retro_insur = trim(replace(shel_retro_insur, "_", ""))
		EMReadScreen shel_retro_insur_ver, 2, 14, 48
		EMReadScreen shel_prosp_insur, 8, 14, 56
			shel_prosp_insur = trim(replace(shel_prosp_insur, "_", ""))
		EMReadScreen shel_prosp_insur_ver, 2, 14, 67
		EMReadScreen shel_retro_taxes, 8, 15, 37
			shel_retro_taxes = trim(replace(shel_retro_taxes, "_", ""))
		EMReadScreen shel_retro_taxes_ver, 2, 15, 48
		EMReadScreen shel_prosp_taxes, 8, 15, 56
			shel_prosp_taxes = trim(replace(shel_prosp_taxes, "_", ""))
		EMReadScreen shel_prosp_taxes_ver, 2, 15, 67
		EMReadScreen shel_retro_room, 8, 16, 37
			shel_retro_room = trim(replace(shel_retro_room, "_", ""))
		EMReadScreen shel_retro_room_ver, 2, 16, 48
		EMReadScreen shel_prosp_room, 8, 16, 56
			shel_prosp_room = trim(replace(shel_prosp_room, "_", ""))
		EMReadScreen shel_prosp_room_ver, 2, 16, 67
		EMReadScreen shel_retro_garage, 8, 17, 37
			shel_retro_garage = trim(replace(shel_retro_garage, "_", ""))
		EMReadScreen shel_retro_garage_ver, 2, 17, 48
		EMReadScreen shel_prosp_garage, 8, 17, 56
			shel_prosp_garage = trim(replace(shel_prosp_garage, "_", ""))
		EMReadScreen shel_prosp_garage_ver, 2, 17, 67
		EMReadScreen shel_retro_subsidy, 8, 18, 37
			shel_retro_subsidy = trim(replace(shel_retro_subsidy, "_", ""))
		EMReadScreen shel_retro_subsidy_ver, 2, 18, 48
		EMReadScreen shel_prosp_subsidy, 8, 18, 56
			shel_prosp_subsidy = trim(replace(shel_prosp_subsidy, "_", ""))
		EMReadScreen shel_prosp_subsidy_ver, 2, 18, 67

		IF shel_subsidized <> "_" 			THEN CALL check_for_data_validation(shel_row, excel_col, shel_subsidized, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_shared  <> "_" 				THEN CALL check_for_data_validation(shel_row + 1, excel_col, shel_shared, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_paid_to  <> "" 				THEN objExcel.Cells(shel_row + 2, excel_col).Value = shel_paid_to
		IF shel_retro_rent <> "" 			THEN objExcel.Cells(shel_row + 3, excel_col).Value = shel_retro_rent
		IF shel_retro_rent_ver <> "__" 		THEN CALL check_for_data_validation(shel_row + 4, excel_col, shel_retro_rent_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_rent <> "" 			THEN objExcel.Cells(shel_row + 5, excel_col).Value = shel_prosp_rent
		IF shel_prosp_rent_ver <> "__" 		THEN CALL check_for_data_validation(shel_row + 6, excel_col, shel_prosp_rent_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_lot_rent <> "" 		THEN objExcel.Cells(shel_row + 7, excel_col).Value = shel_retro_lot_rent
		IF shel_retro_lot_rent_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 8, excel_col, shel_retro_lot_rent_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_lot_rent <> "" 		THEN objExcel.Cells(shel_row + 9, excel_col).Value = shel_prosp_lot_rent
		IF shel_prosp_lot_rent_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 10, excel_col, shel_prosp_lot_rent_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_mortgage <> "" 		THEN objExcel.Cells(shel_row + 11, excel_col).Value = shel_retro_mortgage
		IF shel_retro_mortgage_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 12, excel_col, shel_retro_mortgage_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_mortgage <> "" 		THEN objExcel.Cells(shel_row + 13, excel_col).Value = shel_prosp_mortgage
		IF shel_prosp_mortgage_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 14, excel_col, shel_prosp_mortgage_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_insur <> "" 			THEN objExcel.Cells(shel_row + 15, excel_col).Value = shel_retro_insur
		IF shel_retro_insur_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 16, excel_col, shel_retro_insur_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_insur <> "" 			THEN objExcel.Cells(shel_row + 17, excel_col).Value = shel_prosp_insur
		IF shel_prosp_insur_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 18, excel_col, shel_prosp_insur_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_taxes <> "" 			THEN objExcel.Cells(shel_row + 19, excel_col).Value = shel_retro_taxes
		IF shel_retro_taxes_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 20, excel_col, shel_retro_taxes_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_taxes <> "" 			THEN objExcel.Cells(shel_row + 21, excel_col).Value = shel_prosp_taxes
		IF shel_prosp_taxes_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 22, excel_col, shel_prosp_taxes_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_room <> "" 			THEN objExcel.Cells(shel_row + 23, excel_col).Value = shel_retro_room
		IF shel_retro_room_ver <> "__" 		THEN CALL check_for_data_validation(shel_row + 24, excel_col, shel_retro_room_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_room <> "" 			THEN objExcel.Cells(shel_row + 25, excel_col).Value = shel_prosp_room
		IF shel_prosp_room_ver <> "__" 		THEN CALL check_for_data_validation(shel_row + 26, excel_col, shel_prosp_room_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_garage <> "" 			THEN objExcel.Cells(shel_row + 27, excel_col).Value = shel_retro_garage
		IF shel_retro_garage_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 28, excel_col, shel_retro_garage_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_garage <> "" 			THEN objExcel.Cells(shel_row + 29, excel_col).Value = shel_prosp_garage
		IF shel_prosp_garage_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 30, excel_col, shel_prosp_garage_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_retro_subsidy <> "" 		THEN objExcel.Cells(shel_row + 31, excel_col).Value = shel_retro_subsidy
		IF shel_retro_subsidy_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 32, excel_col, shel_retro_subsidy_ver, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF shel_prosp_subsidy <> "" 		THEN objExcel.Cells(shel_row + 33, excel_col).Value = shel_prosp_subsidy
		IF shel_prosp_subsidy_ver <> "__" 	THEN CALL check_for_data_validation(shel_row + 34, excel_col, shel_prosp_subsidy_ver, objExcel, objWorkbook, objTemplate, objNewSheet)

	END IF
	excel_col = excel_col + 1
NEXT

'SIBL
CALL navigate_to_MAXIS_screen("STAT", "SIBL")
EMReadScreen num_of_sibl, 1, 2, 78
IF num_of_sibl <> "0" THEN 
	EMReadScreen sibl_group1, 38, 7, 39
		sibl_group1 = replace(sibl_group1, "  ", ",")
		sibl_group1 = replace(sibl_group1, ",__", "")
		sibl_group1 = trim(sibl_group1)
	EMReadScreen sibl_group2, 38, 8, 39
		sibl_group2 = replace(sibl_group2, "  ", ",")	
		sibl_group2 = replace(sibl_group2, ",__", "")	
		sibl_group2 = trim(sibl_group2)
	EMReadScreen sibl_group3, 38, 9, 39
		sibl_group3 = replace(sibl_group3, "  ", ",")
		sibl_group3 = replace(sibl_group3, ",__", "")	
		sibl_group3 = trim(sibl_group3)
	
	IF sibl_group1 <> "__" THEN objExcel.Cells(sibl_row, 3).Value = sibl_group1
	IF sibl_group2 <> "__" THEN objExcel.Cells(sibl_row + 1, 3).Value = sibl_group2
	IF sibl_group3 <> "__" THEN objExcel.Cells(sibl_row + 2, 3).Value = sibl_group3
END IF

'SPON
CALL navigate_to_MAXIS_screen("STAT", "SPON")
excel_col = 3
FOR EACH client IN client_array
	EMReadScreen num_of_spon, 1, 2, 78
	IF num_of_spon <> "0" THEN 
		EMReadScreen spon_type, 2, 6, 38
		EMReadScreen spon_verif, 1, 6, 62
		EMReadScreen spon_name, 20, 8, 38
			spon_name = replace(spon_name, "_", "")
			spon_name = trim(spon_name)
		EMReadScreen spon_state, 2, 10, 62
	
		CALL check_for_data_validation(spon_row, excel_col, spon_type, objExcel, objWorkbook, objTemplate, objNewSheet)
		CALL check_for_data_validation(spon_row + 1, excel_col, spon_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		objExcel.Cells(spon_row + 3, excel_col).Value = spon_name
		objExcel.Cells(spon_row + 4, excel_col).Value = spon_state
	END IF
	excel_col = excel_col + 1
NEXT
 
'STEC
CALL navigate_to_MAXIS_screen("STAT", "STEC")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_stec, 1, 2, 78
	IF num_of_stec <> "0" THEN 
		EMReadScreen stec1_type, 2, 8, 25
		IF stec1_type <> "__" THEN 
			EMReadScreen stec1_amt, 8, 8, 31
				stec1_amt = trim(stec1_amt)
				stec1_amt = replace(stec1_amt, "_", "")
			EMReadScreen stec1_thru_dts, 12, 8, 41
				stec1_thru_dts = replace(stec1_thru_dts, "  ", "-")
				stec1_thru_dts = replace(stec1_thru_dts, " ", "/")
			EMReadScreen stec1_verif, 1, 8, 55
			EMReadScreen stec1_earmarked, 8, 8, 59
				stec1_earmarked = trim(stec1_earmarked)
				stec1_earmarked = replace(stec1_earmarked, "_", "")
			EMReadScreen stec1_ear_mos, 12, 8, 69
				stec1_ear_mos = replace(stec1_ear_mos, "  ", "-")
				stec1_ear_mos = replace(stec1_ear_mos, " ", "/")
			
			CALL check_for_data_validation(stec_row, excel_col, stec1_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(stec_row + 1, excel_col).Value = stec1_amt
			IF stec1_thru_dts <> "__/__-__/__" THEN objExcel.Cells(stec_row + 2, excel_col).Value = stec1_thru_dts
			CALL check_for_data_validation(stec_row + 3, excel_col, stec1_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(stec_row + 4, excel_col).Value = stec1_earmarked
			IF stec1_ear_mos <> "__/__-__/__" THEN objExcel.Cells(stec_row + 5, excel_col).Value stec1_ear_mos
			
			EMReadScreen stec2_type, 2, 9, 25
			IF stec2_type <> "__" THEN 
				EMReadScreen stec2_amt, 8, 8, 31
					stec2_amt = trim(stec2_amt)
					stec2_amt = replace(stec2_amt, "_", "")
				EMReadScreen stec2_thru_dts, 12, 8, 41
					stec2_thru_dts = replace(stec2_thru_dts, "  ", "-")
					stec2_thru_dts = replace(stec2_thru_dts, " ", "/")
				EMReadScreen stec2_verif, 1, 8, 55
				EMReadScreen stec2_earmarked, 8, 8, 59
					stec2_earmarked = trim(stec2_earmarked)
					stec2_earmarked = replace(stec2_earmarked, "_", "")
				EMReadScreen stec2_ear_mos, 12, 8, 69
					stec2_ear_mos = replace(stec2_ear_mos, "  ", "-")
					stec2_ear_mos = replace(stec2_ear_mos, " ", "/")
				
				CALL check_for_data_validation(stec_row + 6, excel_col, stec2_type, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(stec_row + 7, excel_col).Value = stec2_amt
				IF stec2_thru_dts <> "__/__-__/__" THEN objExcel.Cells(stec_row + 8, excel_col).Value = stec2_thru_dts
				CALL check_for_data_validation(stec_row + 9, excel_col, stec2_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(stec_row + 10, excel_col).Value = stec2_earmarked
				IF stec2_ear_mos <> "__/__-__/__" THEN objExcel.Cells(stec_row + 11, excel_col).Value stec2_ear_mos			
			END IF
		END IF
	END IF
	excel_col = excel_col + 1
NEXT

'STIN
CALL navigate_to_MAXIS_screen("STAT", "STIN")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_stin, 1, 2, 78
	IF num_of_stin <> "0" THEN
		EMReadScreen stin_type1, 2, 8, 27
		IF stin_type1 <> "__" THEN 
			EMReadScreen stin_amt1, 8, 8, 34
				stin_amt1 = trim(stin_amt1)
				stin_amt1 = replace(stin_amt1, "_", "")
			EMReadScreen stin_avail1, 8, 8, 46
			EMReadScreen stin_thru_dts1, 14, 8, 58
				stin_thru_dts1 = replace(stin_thru_dts1, "    ", "-")
				stin_thru_dts1 = replace(stin_thru_dts1, " ", "/")
			EMReadScreen stin_verif1, 1, 8, 76
			
			CALL check_for_data_validation(stin_row, excel_col, stin_type1, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(stin_row + 1, excel_col).Value = stin_amt1
			IF stin_avail1 <> "__ __ __" THEN objExcel.Cells(stin_row + 2, excel_col).Value = replace(stin_avail1, " ", "/")
			IF stin_thru_dts1 <> "__/__-__/__" THEN objExcel.Cells(stin_row + 3, excel_col).Value = stin_thru_dts1
			IF stin_verif1 <> "_" THEN CALL check_for_data_validation(stin_row + 4, excel_col, stin_verif1, objExcel, objWorkbook, objTemplate, objNewSheet)
			
			EMReadScreen stin_type2, 2, 9, 27
			IF stin_type2 <> "__" THEN 
				EMReadScreen stin_amt2, 8, 9, 34
					stin_amt2 = trim(stin_amt2)
					stin_amt2 = replace(stin_amt2, "_", "")
				EMReadScreen stin_avail2, 8, 9, 46
				EMReadScreen stin_thru_dts2, 14, 9, 58
					stin_thru_dts2 = replace(stin_thru_dts2, "    ", "-")
					stin_thru_dts2 = replace(stin_thru_dts2, " ", "/")
				EMReadScreen stin_verif2, 1, 9, 76
				
				CALL check_for_data_validation(stin_row + 5, excel_col, stin_type2, objExcel, objWorkbook, objTemplate, objNewSheet)
				objExcel.Cells(stin_row + 6, excel_col).Value = stin_amt2
				IF stin_avail2 <> "__ __ __" THEN objExcel.Cells(stin_row + 7, excel_col).Value = replace(stin_avail2, " ", "/")
				IF stin_thru_dts2 <> "__/__-__/__" THEN objExcel.Cells(stin_row + 8, excel_col).Value = stin_thru_dts2
				IF stin_verif2 <> "_" THEN CALL check_for_data_validation(stin_row + 9, excel_col, stin_verif2, objExcel, objWorkbook, objTemplate, objNewSheet)
			END IF
		END IF
	END IF
	excel_col = excel_col + 1
NEXT

'STWK
CALL navigate_to_MAXIS_screen("STAT", "STWK")
excel_col = 3
FOR EACH client IN client_array
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_stwk, 1, 2, 78
	IF num_of_stwk <> "0" THEN 
		EMReadScreen stwk_employer, 30, 6, 46
			stwk_employer = trim(stwk_employer)
			stwk_employer = replace(stwk_employer, "_", "")
		EMReadScreen stwk_stop_work_dt, 8, 7, 46
		EMReadScreen stwk_verif, 1, 7, 63
		EMReadScreen stwk_inc_stop_dt, 8, 8, 46
		EMReadScreen stwk_refuse_emp, 1, 8, 78
		EMReadScreen stwk_vol_quit, 1, 10, 46
		EMReadScreen stwk_refuse_emp_date, 8, 10, 72
		EMReadScreen stwk_good_cause_cash, 1, 12, 52
		EMReadScreen stwk_good_cause_grh, 1, 12, 60
		EMReadScreen stwk_good_cause_snap, 1, 12, 67
		EMReadScreen stwk_fs_pwe, 1, 14, 46
		EMReadScreen stwk_maepd, 1, 16, 46
		
		IF stwk_employer <> "" 					THEN objExcel.Cells(stwk_row, excel_col).Value = stwk_employer
		IF stwk_stop_work_dt <> "__ __ __"		THEN objExcel.Cells(stwk_row + 1, excel_col).Value = replace(stwk_stop_work_dt, " ", "/")
		IF stwk_verif <> "_" 					THEN CALL check_for_data_validation(stwk_row + 2, excel_col, stwk_verif, objExcel, objWorkbook, objTemplate, objNewSheet)	
		IF stwk_inc_stop_dt <> "__ __ __"		THEN objExcel.Cells(stwk_row + 3, excel_col).Value = replace(stwk_inc_stop_dt, " ", "/")
		IF stwk_refuse_emp <> "_" 				THEN CALL check_for_data_validation(stwk_row + 4, excel_col, stwk_refuse_emp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF stwk_vol_quit <> "_" 				THEN CALL check_for_data_validation(stwk_row + 5, excel_col, stwk_vol_quit, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF stwk_refuse_emp_date <> "__ __ __"	THEN objExcel.Cells(stwk_row + 6, excel_col).Value = replace(stwk_refuse_emp_date, " ", "/")
		IF stwk_good_cause_cash <> "_"			THEN CALL check_for_data_validation(stwk_row + 7, excel_col, stwk_good_cause_cash, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF stwk_good_cause_grh <> "_" 			THEN CALL check_for_data_validation(stwk_row + 8, excel_col, stwk_good_cause_grh, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF stwk_good_cause_snap <> "_" 			THEN CALL check_for_data_validation(stwk_row + 9, excel_col, stwk_good_cause_snap, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF stwk_fs_pwe <> "_" 					THEN CALL check_for_data_validation(stwk_row + 10, excel_col, stwk_fs_pwe, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF stwk_maepd <> "_" 					THEN CALL check_for_data_validation(stwk_row + 11, excel_col, stwk_maepd, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'UNEA1, UNEA2, UNEA3
CALL navigate_to_MAXIS_screen("STAT", "UNEA")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_unea, 1, 2, 78
	IF num_of_unea <> "0" THEN
		num_of_active_unea = 0
		DO
			EMReadScreen unea_end_date, 8, 7, 68
			IF unea_end_date = "__ __ __" THEN num_of_active_unea = num_of_active_unea + 1
			transmit
			EMReadScreen enter_a_valid, 13, 24, 2
		LOOP UNTIL enter_a_valid = "ENTER A VALID"	
		
		IF num_of_active_unea >= 1 THEN
			CALL write_value_and_transmit("01", 20, 79)
			DO
				EMReadScreen unea_end_date, 8, 7, 68
				IF unea_end_date = "__ __ __" THEN EXIT DO
				transmit
			LOOP
			EMReadScreen unea1_type, 2, 5, 37
				IF unea1_type = "01" OR unea1_type = "02" OR unea1_type = "03" THEN 
					EMReadScreen unea1_suffix, 3, 6, 46
					objExcel.Cells(unea1_row + 2, excel_col).Value = unea1_suffix
				END IF
			EMReadScreen unea1_verif, 1, 5, 65
			EMReadScreen unea1_inc_start, 8, 7, 37
				unea1_inc_start = replace(unea1_inc_start, " ", "/")
			'Grabbing the pay frequency off the PIC...if nothing is there, the script will assume once per month
			IF snap_case = TRUE THEN
				CALL write_value_and_transmit("X", 10, 26)
				EMReadScreen unea1_pay_freq, 1, 5, 64
				EMReadScreen unea1_pay_amt, 8, 17, 56
					unea1_pay_amt = trim(unea1_pay_amt)
				transmit
			END IF
			IF snap_case = FALSE AND health_care_case = TRUE THEN 
				CALL write_value_and_transmit("X", 6, 56)
				EMReadscreen unea1_pay_freq, 1, 10, 63
				IF unea1_pay_freq = "_" THEN unea1_pay_freq = "1"
				EMReadScreen unea1_pay_amt, 8, 9, 65
					unea1_pay_amt = trim(replace(unea1_pay_amt, "_", ""))
				transmit
			END IF
			IF snap_case = FALSE AND health_care_case = FALSE AND cash_case = TRUE THEN 
				'assuming pay frequency is once monthly if we cannot find it anywhere
				unea1_pay_freq = "1"
				EMReadScreen unea1_pay_amt, 8, 18, 68
				unea1_pay_amt = trim(unea1_pay_amt)
			END IF
			
			CALL check_for_data_validation(unea1_row, excel_col, unea1_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(unea1_row + 1, excel_col, unea1_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(unea1_row + 3, excel_col).Value = unea1_inc_start
			objExcel.Cells(unea1_row + 4, excel_col).Value = unea1_pay_freq
			objExcel.Cells(unea1_row + 5, excel_col).Value = unea1_pay_amt			
		END IF
		
		IF num_of_active_unea >= 2 THEN
			transmit
			DO
				EMReadScreen unea_end_date, 8, 7, 68
				IF unea_end_date = "__ __ __" THEN EXIT DO
				transmit
			LOOP
			EMReadScreen unea2_type, 2, 5, 37
				IF unea2_type = "01" OR unea2_type = "02" OR unea2_type = "03" THEN 
					EMReadScreen unea2_suffix, 3, 6, 46
					objExcel.Cells(unea2_row + 2, excel_col).Value = unea2_suffix
				END IF			
			EMReadScreen unea2_verif, 1, 5, 65
			EMReadScreen unea2_inc_start, 8, 7, 37
				unea2_inc_start = replace(unea2_inc_start, " ", "/")
			'Grabbing the pay frequency off the PIC...if nothing is there, the script will assume once per month
			IF snap_case = TRUE THEN
				CALL write_value_and_transmit("X", 10, 26)
				EMReadScreen unea2_pay_freq, 1, 5, 64
				EMReadScreen unea2_pay_amt, 8, 17, 56
					unea2_pay_amt = trim(unea2_pay_amt)
				transmit
			END IF
			IF snap_case = FALSE AND health_care_case = TRUE THEN 
				CALL write_value_and_transmit("X", 6, 56)
				EMReadscreen unea2_pay_freq, 1, 10, 63
				IF unea2_pay_freq = "_" THEN unea2_pay_freq = "1"
				EMReadScreen unea2_pay_amt, 8, 9, 65
					unea2_pay_amt = trim(replace(unea2_pay_amt, "_", ""))
				transmit
			END IF
			IF snap_case = FALSE AND health_care_case = FALSE AND cash_case = TRUE THEN 
				'assuming pay frequency is once monthly if we cannot find it anywhere
				unea2_pay_freq = "1"
				EMReadScreen unea2_pay_amt, 8, 18, 68
				unea2_pay_amt = trim(unea2_pay_amt)
			END IF
			
			CALL check_for_data_validation(unea2_row, excel_col, unea2_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(unea2_row + 1, excel_col, unea2_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(unea2_row + 3, excel_col).Value = unea2_inc_start
			objExcel.Cells(unea2_row + 4, excel_col).Value = unea2_pay_freq
			objExcel.Cells(unea2_row + 5, excel_col).Value = unea2_pay_amt	
		END IF		
		
		IF num_of_active_unea >= 3 THEN
			transmit
			DO
				EMReadScreen unea_end_date, 8, 7, 68
				IF unea_end_date = "__ __ __" THEN EXIT DO
				transmit
			LOOP
			EMReadScreen unea3_type, 2, 5, 37
				IF unea3_type = "01" OR unea3_type = "02" OR unea3_type = "03" THEN 
					EMReadScreen unea3_suffix, 3, 6, 46
					objExcel.Cells(unea3_row + 2, excel_col).Value = unea3_suffix
				END IF			
			EMReadScreen unea3_verif, 1, 5, 65
			EMReadScreen unea3_inc_start, 8, 7, 37
				unea3_inc_start = replace(unea3_inc_start, " ", "/")
			'Grabbing the pay frequency off the PIC...if nothing is there, the script will assume once per month
			IF snap_case = TRUE THEN
				CALL write_value_and_transmit("X", 10, 26)
				EMReadScreen unea3_pay_freq, 1, 5, 64
				EMReadScreen unea3_pay_amt, 8, 17, 56
					unea3_pay_amt = trim(unea3_pay_amt)
				transmit
			END IF
			IF snap_case = FALSE AND health_care_case = TRUE THEN 
				CALL write_value_and_transmit("X", 6, 56)
				EMReadscreen unea3_pay_freq, 1, 10, 63
				IF unea3_pay_freq = "_" THEN unea3_pay_freq = "1"
				EMReadScreen unea3_pay_amt, 8, 9, 65
					unea3_pay_amt = trim(replace(unea3_pay_amt, "_", ""))
				transmit
			END IF
			IF snap_case = FALSE AND health_care_case = FALSE AND cash_case = TRUE THEN 
				'assuming pay frequency is once monthly if we cannot find it anywhere
				unea3_pay_freq = "1"
				EMReadScreen unea3_pay_amt, 8, 18, 68
				unea3_pay_amt = trim(unea3_pay_amt)
			END IF
			
			CALL check_for_data_validation(unea3_row, excel_col, unea3_type, objExcel, objWorkbook, objTemplate, objNewSheet)
			CALL check_for_data_validation(unea3_row + 1, excel_col, unea3_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
			objExcel.Cells(unea3_row + 3, excel_col).Value = unea3_inc_start
			objExcel.Cells(unea3_row + 4, excel_col).Value = unea3_pay_freq
			objExcel.Cells(unea3_row + 5, excel_col).Value = unea3_pay_amt		
		END IF				
	END IF
	excel_col = excel_col + 1
NEXT

'WKEX
CALL navigate_to_MAXIS_screen("STAT", "WKEX")
excel_col = 3
FOR EACH client IN client_array
	EMWriteScreen client, 20, 76
	CALL write_value_and_transmit("01", 20, 79)
	EMReadScreen num_of_wkex, 1, 2, 78
	IF num_of_wkex <> "0" THEN 
		'reading the first WKEX panel
		EMReadScreen wkex_program, 2, 5, 33
		EMReadScreen wkex_fed_tax_retro, 8, 7, 43
			wkex_fed_tax_retro = trim(replace(wkex_fed_tax_retro, "_", ""))
		EMReadScreen wkex_fed_tax_prosp, 8, 7, 57
			wkex_fed_tax_prosp = trim(replace(wkex_fed_tax_prosp, "_", ""))
		EMReadScreen wkex_fed_tax_verif, 1, 7, 69
		EMReadScreen wkex_state_tax_retro, 8, 8, 43
			wkex_state_tax_retro = trim(replace(wkex_state_tax_retro, "_", ""))
		EMReadScreen wkex_state_tax_prosp, 8, 8, 57
			wkex_state_tax_prosp = trim(replace(wkex_state_tax_prosp, "_", ""))
		EMReadScreen wkex_state_tax_verif, 1, 8, 69
		EMReadScreen wkex_fica_retro, 8, 9, 43
			wkex_fica_retro = trim(replace(wkex_fica_retro, "_", ""))
		EMReadScreen wkex_fica_prosp, 8, 9, 57
			wkex_fica_prosp = trim(replace(wkex_fica_prosp, "_", ""))
		EMReadScreen wkex_fica_verif, 1, 9, 69
		EMReadScreen wkex_trans_retro, 8, 10, 43
			wkex_trans_retro = trim(replace(wkex_trans_retro, "_", ""))
		EMReadScreen wkex_trans_prosp, 8, 10, 57
			wkex_trans_prosp = trim(replace(wkex_trans_prosp, "_", ""))
		EMReadScreen wkex_trans_verif, 1, 10, 69
		EMReadScreen wkex_trans_imp, 1, 10, 75
		EMReadScreen wkex_meals_retro, 8, 11, 43
			wkex_meals_retro = trim(replace(wkex_meals_retro, "_", ""))
		EMReadScreen wkex_meals_prosp, 8, 11, 57
			wkex_meals_prosp = trim(replace(wkex_meals_prosp, "_", ""))
		EMReadScreen wkex_meals_verif, 1, 11, 69
		EMReadScreen wkex_meals_imp, 1, 11, 75
		EMReadScreen wkex_uniforms_retro, 8, 12, 43
			wkex_uniforms_retro = trim(replace(wkex_uniforms_retro, "_", ""))
		EMReadScreen wkex_uniforms_prosp, 8, 12, 57
			wkex_uniforms_prosp = trim(replace(wkex_uniforms_prosp, "_", ""))
		EMReadScreen wkex_uniforms_verif, 1, 12, 69
		EMReadScreen wkex_uniforms_imp, 1, 12, 75
		EMReadScreen wkex_tools_retro, 8, 13, 43
			wkex_tools_retro = trim(replace(wkex_tools_retro, "_", ""))
		EMReadScreen wkex_tools_prosp, 8, 13, 57
			wkex_tools_prosp = trim(replace(wkex_tools_prosp, "_", ""))
		EMReadScreen wkex_tools_verif, 1, 13, 69
		EMReadScreen wkex_tools_imp, 1, 13, 75
		'...there's a lot to read...
		EMReadScreen wkex_dues_retro, 8, 14, 43
			wkex_dues_retro = trim(replace(wkex_dues_retro, "_", ""))
		EMReadScreen wkex_dues_prosp, 8, 14, 57
			wkex_dues_prosp = trim(replace(wkex_dues_prosp, "_", ""))
		EMReadScreen wkex_dues_verif, 1, 14, 69
		EMReadScreen wkex_dues_imp, 1, 14, 75
		EMReadScreen wkex_other_retro, 8, 15, 43
			wkex_other_retro = trim(replace(wkex_other_retro, "_", ""))
		EMReadScreen wkex_other_prosp, 8, 15, 57
			wkex_other_prosp = trim(replace(wkex_other_prosp, "_", ""))
		EMReadScreen wkex_other_verif, 1, 15, 69
		EMReadScreen wkex_other_imp, 1, 15, 75
		'going in to the HC expense est
		CALL write_value_and_transmit("X", 18, 57)
		EMReadScreen wkex_hc_fed, 8, 8, 36
			wkex_hc_fed = trim(replace(wkex_hc_fed, "_", ""))
		EMReadScreen wkex_hc_state, 8, 9, 36
			wkex_hc_state = trim(replace(wkex_hc_state, "_", ""))
		EMReadScreen wkex_hc_fica, 8, 10, 36
			wkex_hc_fica = trim(replace(wkex_hc_fica, "_", ""))
		EMReadScreen wkex_hc_trans, 8, 11, 36
			wkex_hc_trans = trim(replace(wkex_hc_trans, "_", ""))
		EMReadScreen wkex_hc_trans_imp, 1, 11, 51
		EMReadScreen wkex_hc_meals, 8, 12, 36
			wkex_hc_trans = trim(replace(wkex_hc_trans, "_", ""))
		EMReadScreen wkex_hc_meals_imp, 1, 12, 51
		EMReadScreen wkex_hc_unif, 8, 13, 36
			wkex_hc_unif = trim(replace(wkex_hc_unif, "_", ""))
		EMReadScreen wkex_hc_unif_imp, 1, 13, 51
		EMReadScreen wkex_hc_tool, 8, 14, 36
			wkex_hc_tool = trim(replace(wkex_hc_tool, "_", ""))
		EMReadScreen wkex_hc_tool_imp, 1, 14, 51
		EMReadScreen wkex_hc_dues, 8, 15, 36
			wkex_hc_dues = trim(replace(wkex_hc_dues, "_", ""))
		EMReadScreen wkex_hc_dues_imp, 1, 15, 51
		EMReadScreen wkex_hc_other, 8, 16, 36
			wkex_hc_other = trim(replace(wkex_hc_other, "_", ""))
		EMReadScreen wkex_hc_other_imp, 1, 15, 51
		'and leaving...
		PF3
		'and finally done reading from WKEX
		
		'writing
		IF wkex_program <> "__" 			THEN CALL check_for_data_validation(wkex_row, excel_col, wkex_program, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_fed_tax_retro <> ""			THEN objExcel.Cells(wkex_row + 1, excel_col).Value = wkex_fed_tax_retro
		IF wkex_fed_tax_prosp <> ""			THEN objExcel.Cells(wkex_row + 2, excel_col).Value = wkex_fed_tax_prosp
		IF wkex_fed_tax_verif <> "_"		THEN CALL check_for_data_validation(wkex_row + 3, excel_col, wkex_fed_tax_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_state_tax_retro <> ""		THEN objExcel.Cells(wkex_row + 4, excel_col).Value = wkex_state_tax_retro
		IF wkex_state_tax_prosp <> "" 		THEN objExcel.Cells(wkex_row + 5, excel_col).Value = wkex_state_tax_prosp
		IF wkex_state_tax_verif <> "_"		THEN CALL check_for_data_validation(wkex_row + 6, excel_col, wkex_state_tax_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_fica_retro <> ""			THEN objExcel.Cells(wkex_row + 7, excel_col).Value = wkex_fica_retro
		IF wkex_fica_prosp <> ""			THEN objExcel.Cells(wkex_row + 8, excel_col).Value = wkex_fica_prosp
		IF wkex_fica_verif <> "_"			THEN CALL check_for_data_validation(wkex_row + 9, excel_col, wkex_fica_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_trans_retro <> ""			THEN objExcel.Cells(wkex_row + 10, excel_col).Value = wkex_trans_retro
		IF wkex_trans_prosp <> ""			THEN objExcel.Cells(wkex_row + 11, excel_col).Value = wkex_trans_prosp
		IF wkex_trans_verif	<> "_"			THEN CALL check_for_data_validation(wkex_row + 12, excel_col, wkex_trans_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_trans_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 13, excel_col, wkex_trans_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_meals_retro <> ""			THEN objExcel.Cells(wkex_row + 14, excel_col).Value = wkex_meals_retro
		IF wkex_meals_prosp <> ""			THEN objExcel.Cells(wkex_row + 15, excel_col).Value = wkex_meals_prosp
		IF wkex_meals_verif <> "_"			THEN CALL check_for_data_validation(wkex_row + 16, excel_col, wkex_meals_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_meals_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 17, excel_col, wkex_meals_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_uniforms_retro <> ""		THEN objExcel.Cells(wkex_row + 18, excel_col).Value = wkex_uniforms_retro
		IF wkex_uniforms_prosp <> ""		THEN objExcel.Cells(wkex_row + 19, excel_col).Value = wkex_uniforms_prosp
		IF wkex_uniforms_verif <> "_"		THEN CALL check_for_data_validation(wkex_row + 20, excel_col, wkex_uniforms_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_uniforms_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 21, excel_col, wkex_uniforms_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_tools_retro <> ""			THEN objExcel.Cells(wkex_row + 22, excel_col).Value = wkex_tools_retro
		IF wkex_tools_prosp <> ""			THEN objExcel.Cells(wkex_row + 23, excel_col).Value = wkex_fed_tax_prosp
		IF wkex_tools_verif <> "_"			THEN CALL check_for_data_validation(wkex_row + 24, excel_col, wkex_tools_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_tools_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 25, excel_col, wkex_tools_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_dues_retro <> ""			THEN objExcel.Cells(wkex_row + 26, excel_col).Value = wkex_dues_retro
		IF wkex_dues_prosp <> ""			THEN objExcel.Cells(wkex_row + 27, excel_col).Value = wkex_dues_prosp
		IF wkex_dues_verif <> "_"			THEN CALL check_for_data_validation(wkex_row + 28, excel_col, wkex_dues_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_dues_imp <> "_"				THEN CALL check_for_data_validation(wkex_row + 29, excel_col, wkex_dues_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_other_retro <> ""			THEN objExcel.Cells(wkex_row + 30, excel_col).Value = wkex_other_retro
		IF wkex_other_prosp <> ""			THEN objExcel.Cells(wkex_row + 31, excel_col).Value = wkex_other_prosp
		IF wkex_other_verif <> "_"			THEN CALL check_for_data_validation(wkex_row + 32, excel_col, wkex_other_verif, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_other_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 33, excel_col, wkex_other_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_hc_fed <> "" 				THEN objExcel.Cells(wkex_row + 34, excel_col).Value = wkex_hc_fed
		IF wkex_hc_state <> "" 				THEN objExcel.Cells(wkex_row + 35, excel_col).Value = wkex_hc_state
		IF wkex_hc_fica <> "" 				THEN objExcel.Cells(wkex_row + 36, excel_col).Value = wkex_hc_fica
		IF wkex_hc_trans <> "" 				THEN objExcel.Cells(wkex_row + 37, excel_col).Value = wkex_hc_trans
		IF wkex_hc_trans_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 38, excel_col, wkex_hc_trans_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_hc_meals <> "" 				THEN objExcel.Cells(wkex_row + 39, excel_col).Value = wkex_hc_meals
		IF wkex_hc_meals_imp <> "_" 		THEN CALL check_for_data_validation(wkex_row + 40, excel_col, wkex_hc_meals_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_hc_unif <> "" 				THEN objExcel.Cells(wkex_row + 41, excel_col).Value = wkex_hc_unif
		IF wkex_hc_unif_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 42, excel_col, wkex_hc_unif_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_hc_tool <> "" 				THEN objExcel.Cells(wkex_row + 43, excel_col).Value = wkex_hc_tool
		IF wkex_hc_tool_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 44, excel_col, wkex_hc_tool_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_hc_dues <> ""				THEN objExcel.Cells(wkex_row + 45, excel_col).Value = wkex_hc_dues
		IF wkex_hc_dues_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 46, excel_col, wkex_hc_dues_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF wkex_hc_other <> ""				THEN objExcel.Cells(wkex_row + 47, excel_col).Value = wkex_hc_other
		IF wkex_hc_other_imp <> "_"			THEN CALL check_for_data_validation(wkex_row + 48, excel_col, wkex_hc_other_imp, objExcel, objWorkbook, objTemplate, objNewSheet)
		'...and now I'm done with WKEX forever...
	END IF
	excel_col = excel_col + 1
NEXT

'WREG
CALL navigate_to_MAXIS_screen("STAT", "WREG")
excel_col = 3
FOR EACH client IN client_array
	'Navigating to WREG for each ref num
	CALL write_value_and_transmit(client, 20, 76)
	EMReadScreen num_of_wreg, 1, 2, 78
	IF num_of_wreg <> "0" THEN 	
		STATS_manualtime = STATS_manualtime + 28
		'reading
		EMReadScreen fs_pwe, 1, 6, 68
		EMReadScreen fset_status, 2, 8, 50
		EMReadScreen defer_fset, 1, 8, 80
		EMReadScreen fset_orient_dt, 8, 9, 50
			IF fset_orient_dt <> "__ __ __" THEN fset_orient_dt = replace(fset_orient_dt, " ", "/")
		EMReadScreen fset_sanction_dt, 8, 10, 50
			IF fset_sanction_dt <> "__ __ __" THEN fset_sanction_dt = replace(fset_sanction_dt, " ", "/")
		EMReadScreen num_wreg_sanc, 2, 11, 50
		EMReadScreen abawd_status, 2, 13, 50
		EMReadScreen ga_basis, 2, 15, 50
		
		'writing
		IF fs_pwe <> "_" 					THEN objExcel.Cells(wreg_row, excel_col).Value = fs_pwe
		IF fset_status <> "__" 				THEN CALL check_for_data_validation(wreg_row + 1, excel_col, fset_status, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF defer_fset <> "_" 				THEN objExcel.Cells(wreg_row + 2, excel_col).Value = defer_fset
		IF fset_orient_dt <> "__ __ __" 	THEN objExcel.Cells(wreg_row + 3, excel_col).Value = fset_orient_dt
		IF fset_sanction_dt <> "__ __ __" 	THEN objExcel.Cells(wreg_row + 4, excel_col).Value = fset_sanction_dt
		IF num_wreg_sanc <> "__" 			THEN objExcel.Cells(wreg_row + 5, excel_col).Value = num_wreg_sanc
		IF abawd_status <> "__" 			THEN CALL check_for_data_validation(wreg_row + 6, excel_col, abawd_status, objExcel, objWorkbook, objTemplate, objNewSheet)
		IF ga_basis <> "__" 				THEN CALL check_for_data_validation(wreg_row + 7, excel_col, ga_basis, objExcel, objWorkbook, objTemplate, objNewSheet)
	END IF
	excel_col = excel_col + 1
NEXT

'autofitting the columns with content
FOR i = 3 TO excel_col - 1
	objExcel.Columns(i).Autofit()
NEXT

objWorkbook.SaveAs(training_case_creator_excel_file_path) 
end_time = timer
run_time = end_time - start_time
script_end_procedure("file saved." & vbCr & "manual time = " & STATS_manualtime & vbCr & "script run time = " & run_time)
