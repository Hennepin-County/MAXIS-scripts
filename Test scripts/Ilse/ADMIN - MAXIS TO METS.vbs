'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - MAXIS TO METS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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
call changelog_update("07/16/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog info_dialog, 0, 0, 266, 115, "BULK - MAXIS TO METS"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of HC cases needs to be reviewed for METS conversion."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

BeginDialog excel_row_dialog, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo 
MAXIS_footer_year = CM_yr 

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog info_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Setting the column headers
objExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 4).Value = "MAGI Persons"
objExcel.Cells(1, 5).Value = "Non-MAGI Persons"
objExcel.Cells(1, 6).Value = "# of MAGI"
objExcel.Cells(1, 7).Value = "# of Non-MAGI"
objExcel.Cells(1, 8).Value = "MAGI Household"
objExcel.Cells(1, 9).Value = "Mixed Household"
objExcel.Cells(1, 10).Value = "Non-MAGI Household"
objExcel.Cells(1, 11).Value = "MAGI Review aligned?"
objExcel.Cells(1, 12).Value = "HC ER MONTH"

'And now BOLD because format
FOR i = 1 TO 12
	objExcel.Cells(1, i).Font.Bold = true 
NEXT

do 
    dialog excel_row_dialog
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Going back through and determining MAGI & Non-MAGI
excel_row = excel_row_to_restart

Do
	'Assigning a value to MAXIS_case_number
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	'When the script gets to the end of the list, it exits the do/loop
	If trim(MAXIS_case_number) = "" then exit do

	'Reseting critical values
	MAGI_count = 0
	nonMAGI_count = 0
	magi_clients = ""
	non_magi_clients = ""

	'navigating to ELIG/HC
	CALL navigate_to_MAXIS_screen("ELIG", "HC")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
    If PRIV_check = "PRIV" then
        ObjExcel.Cells(excel_row, 3).Value = "PRIV"
        magi_clients = ""
        EMWriteScreen "________", 18, 43
    Else 
        'setting the row to read on ELIG/HC
        hhmm_row = 8
        DO
	        'reading the hc reference number
	        EMReadScreen hc_ref_num, 2, hhmm_row, 3
	        'looking to see that information is found for that client
	        EMReadScreen hc_information_found, 70, hhmm_row, 3
	        hc_information_found = trim(hc_information_found)
	        EMReadScreen elig_result, 4, hhmm_row, 41
	        EMReadScreen elig_status, 6, hhmm_row, 50
	        '...and if information is found for that row...
	        IF hc_information_found <> "" THEN
	        	'...if the client is eligible and active...
	        	IF elig_result = "ELIG" AND elig_status = "ACTIVE" THEN
	        		'looking for the first character on hc request...
	        		EMReadScreen hc_requested, 1, hhmm_row, 28
	        		'...if the client is active on a medicare savings program...
	        		IF hc_requested = "S" OR hc_requested = "Q" OR hc_requested = "I" THEN 			'IF the HH MEMB is MSP ONLY then they are automatically Budg Mthd B
	        			IF hc_ref_num = "  " THEN
	        				temp_hhmm_row = hhmm_row
	        				DO
	        					EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
	        					IF hc_ref_num = "  " THEN
	        						temp_hhmm_row = temp_hhmm_row - 1
	        					ELSE
	        						EXIT DO
	        					END IF
	        				LOOP
	        			END IF
	        			IF InStr(non_magi_clients, hc_ref_num & ";") = 0 THEN non_magi_clients = non_magi_clients & hc_ref_num & ";"
	        			hhmm_row = hhmm_row + 1
	        		'...otherwise, if the client is active on Medicaid or EMA...
	        		ELSEIF hc_requested = "M" or hc_requested = "E" THEN
	        			'...going in to grab the budget method...
	        			EMWriteScreen "X", hhmm_row, 26
	        			transmit
	        			EMReadScreen budg_mthd, 1, 13, 76
	        			IF budg_mthd = "A" THEN
	        				IF hc_ref_num = "  " THEN
	        					temp_hhmm_row = hhmm_row
	        					DO
	        						EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
	        						IF hc_ref_num = "  " THEN
	        							temp_hhmm_row = temp_hhmm_row - 1
	        						ELSE
	        							EXIT DO
	        						END IF
	        					LOOP
	        				END IF
	        				IF InStr(magi_clients, hc_ref_num & ";") = 0 THEN magi_clients = magi_clients & hc_ref_num & ";"
	        			ELSE
	        				IF hc_ref_num = "  " THEN
	        					temp_hhmm_row = hhmm_row
	        					DO
	        						EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
	        						IF hc_ref_num = "  " THEN
	        							temp_hhmm_row = temp_hhmm_row - 1
	        						ELSE
	        							EXIT DO
	        						END IF
	        					LOOP
	        				END IF
	        				IF InStr(non_magi_clients, hc_ref_num & ";") = 0 THEN non_magi_clients = non_magi_clients & hc_ref_num & ";"
	        			END IF
	        			PF3
	        			hhmm_row = hhmm_row + 1
	        		ELSEIF hc_requested = "N" THEN
	        			hhmm_row = hhmm_row + 1
	        		END IF
	        	ELSE
	        		hhmm_row = hhmm_row + 1
	        	END IF
	        ELSE
		        EXIT DO
            End if 
        LOOP UNTIL hhmm_row = 20 OR hc_ref_num = "  "

	   'Going back to determine if the individual is still MAGI...SSI, Medicare, and MA-EPD disq the person as Non-MAGI
	   IF magi_clients <> "" THEN
	   	magi_peeps = replace(magi_clients & "~~~", ";~~~", "")
	   	magi_peeps = split(magi_peeps, ";")
	   	FOR EACH client IN magi_peeps
	   		CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	   		CALL write_value_and_transmit(client, 20, 76)
	   		EMReadScreen client_age, 2, 8, 76
	   		IF client_age = "  " THEN client_age = 0
	   		'Removing that client from the non-magi list
	   		IF client_age >= 65 THEN
	   			magi_clients = replace(magi_clients, client & ";", "")
	   			IF InStr(non_magi_clients, client & ";") = 0 THEN non_magi_clients = non_magi_clients & client & ";"
	   		ELSE
	   			'Checking DISA for a MA-EPD & SSI
	   			CALL navigate_to_MAXIS_screen("STAT", "DISA")
	   			CALL write_value_and_transmit(client, 20, 76)
	   			EMReadScreen hc_disa_status, 2, 13, 59
	   			IF hc_disa_status = "03" OR hc_disa_status = "04" OR hc_disa_status = "22" THEN
	   				magi_clients = replace(magi_clients, client & ";", "")
	   				IF InStr(non_magi_clients, client & ";") = 0 THEN non_magi_clients = non_magi_clients & client & ";"
	   			ELSE
	   				CALL navigate_to_MAXIS_screen("STAT", "MEDI")
	   				CALL write_value_and_transmit(client, 20, 76)
	   				EMReadScreen medi_ref_num, 15, 6, 44
	   				medi_ref_num = trim(replace(medi_ref_num, "_", ""))
	   				IF medi_ref_num <> "" THEN
	   					magi_clients = replace(magi_clients, client & ";", "")
	   					IF InStr(non_magi_clients, client & ";") = 0 THEN non_magi_clients = non_magi_clients & client & ";"
	   				END IF
	   			END IF
	   		END IF
	   	NEXT
	   END IF

	   'Going back to determine if the individual is still Non-MAGI...SSI, Medicare, and MA-EPD disq the person as Non-MAGI
	   IF non_magi_clients <> "" THEN
	   	non_magi_peeps = replace(non_magi_clients & "~~~", ";~~~", "")
	   	non_magi_peeps = split(non_magi_peeps, ";")
	   	FOR EACH client IN non_magi_peeps
	   		'Checking for SSI, MA-EPD, and Medicare
	   		non_magi = ""
	   		CALL navigate_to_MAXIS_screen("STAT", "MEDI")
	   		CALL write_value_and_transmit(client, 20, 76)
	   		EMReadScreen medi_case_number, 15, 6, 44
	   		medi_case_number = trim(replace(medi_case_number, "_", ""))
	   		IF medi_case_number <> "" THEN non_magi = TRUE
	   		IF non_magi <> TRUE THEN
	   			CALL navigate_to_MAXIS_screen("STAT", "DISA")
	   			CALL write_value_and_transmit(client, 20, 76)
	   			EMReadScreen hc_disa_status, 2, 13, 59
	   			IF hc_disa_status = "03" OR hc_disa_status = "04" OR hc_disa_status = "22" THEN non_magi = TRUE
	   		END IF
	   		IF non_magi <> TRUE THEN
	   			non_magi_clients = replace(non_magi_clients, client & ";", "")
	   			IF InStr(magi_clients, client & ";") = 0 THEN magi_clients = magi_clients & client & ";"
	   		END IF
	   	NEXT
	   END IF

	   'Writing all these ding-dang values to Excel
	   objExcel.Cells(excel_row, 4).Value = magi_clients
	   IF magi_clients <> "" THEN
	   	MAGI_count = UBound(split(magi_clients, ";"))
	   	objExcel.Cells(excel_row, 6).Value = MAGI_count
	   ELSE
	   	objExcel.Cells(excel_row, 6).Value = 0
	   END IF

	   objExcel.Cells(excel_row, 5).Value = non_magi_clients
	   IF non_magi_clients <> "" THEN
	   	nonMAGI_count = UBound(split(non_magi_clients, ";"))
	   	objExcel.Cells(excel_row, 7).Value = nonMAGI_count
	   ELSE
	   	objExcel.Cells(excel_row, 7).Value = 0
	   END IF

	   CALL navigate_to_MAXIS_screen("STAT", "REVW")
	   EMReadScreen revw_does_not_exist, 19, 24, 2
	   IF revw_does_not_exist <> "REVW DOES NOT EXIST" THEN
	   	EMwritescreen "X", 5, 71
	   	Transmit
	   	'Checking to make sure pop up opened
	   	DO
	   		EMReadScreen revw_pop_up_check, 8, 4, 44
	   		EMWaitReady 1, 1
	   	LOOP until revw_pop_up_check = "RENEWALS"
	   	'Reading HC reviews to compare them
	   	EMReadScreen hc_income_renewal, 8, 8, 27
	   	EMReadScreen hc_IA_renewal, 8, 8, 71
	   	EMReadScreen hc_annual_renewal, 8, 9, 27
	   	objExcel.Cells(excel_row, 12).Value = replace(hc_annual_renewal, " ", "/")
	   	IF MAGI_count <> 0 THEN
	   		IF hc_income_renewal = "__ 01 __" THEN hc_compare_renewal = hc_IA_renewal
	   		IF hc_IA_renewal = "__ 01 __" THEN hc_compare_renewal = hc_income_renewal

	   		IF hc_annual_renewal = hc_compare_renewal THEN
	   			objExcel.Cells(excel_row, 11).Value = "Y"
	   		ELSE
	   			objExcel.Cells(excel_row, 11).Value = "Y"
	   		END IF
	   	END IF
	   ELSE
	   	objExcel.Cells(excel_row, 12).Value = "NO REVIEW DATE"
	   END IF

	   IF MAGI_count <> 0 AND nonMAGI_count = 0 THEN
	   	objExcel.Cells(excel_row, 8).Value = "Y"
	   ELSEIF MAGI_count <> 0 AND nonMAGI_count <> 0 THEN
	   	objExcel.Cells(excel_row, 9).Value = "Y"
	   ELSEIF MAGI_count = 0 AND nonMAGI_count <> 0 THEN
	   	objExcel.Cells(excel_row, 10).Value = "Y"
	   END IF
    End if 
    
	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!" & vbCr & vbCr & "The script has finished running.")
