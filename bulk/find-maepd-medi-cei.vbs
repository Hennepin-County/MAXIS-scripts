'Required for statistical purposes===============================================================================
name_of_script = "BULK - FIND MAEPD MEDI CEI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 86                      'manual run time in seconds
STATS_denomination = "C"       						 'C is for each CASE
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'Dialogs---------------------------------------------------------------------------------------------------------------------------------
BeginDialog maepd_dlg, 0, 0, 191, 85, "MA-EPD Reimburseables"
  EditBox 100, 10, 65, 15, x_number
  ButtonGroup ButtonPressed
    OkButton 70, 60, 50, 15
    CancelButton 125, 60, 50, 15
  Text 10, 15, 85, 10, "X Number (7 digit format):"
  Text 10, 30, 175, 10, "This script will check REPT/ACTV on this X number."
EndDialog

'The script----------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

CALL check_for_MAXIS(True)
DO
	err_msg = ""								'err message handling to loop until the user has entered the proper information
	Dialog maepd_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF len(x_number) <> 7 THEN err_msg = err_msg & vbCr & "* The X number must be the full 7 digits."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Resolve for the script to continue."
LOOP UNTIL err_msg = ""

CALL check_for_MAXIS(False)
back_to_SELF

CALL navigate_to_MAXIS_screen("REPT", "ACTV")						'navigating to rept actv for requested user
EMReadScreen current_rept_actv, 7, 21, 13
IF ucase(current_rept_actv) <> ucase(x_number) THEN CALL write_value_and_transmit(x_number, 21, 13)			'making sure that the X# was written correctly sometimes there are issues with lower case worker numbers

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
objExcel.Cells(1, 1).Value = "CASE NUMBER"						'creating columns to store the information
objExcel.Cells(1, 2).Value = "CLIENT NAME"
objExcel.Cells(1, 3).Value = "NEXT REVW DT"
objExcel.Cells(1, 4).Value = "REIMBURSEMENT ELIG?"

'setting variables for first run through
rept_row = 7
excel_row = 2
DO
	EMReadScreen last_page, 21, 24, 2											'checking to see if this is the last page, if it is the loop can end.
	DO
		EMReadScreen MAXIS_case_number, 8, rept_row, 12						'reading the case numbers from rept/actv
		MAXIS_case_number = trim(MAXIS_case_number)
		EMReadScreen hc_case, 1, rept_row, 64
		IF MAXIS_case_number <> "" AND hc_case <> " " THEN					'checking for HC cases
			objExcel.Cells(excel_row, 1).Value = MAXIS_case_number			'adding read variables to the spreadsheet
			EMReadScreen client_name, 21, rept_row, 21						'grabbing client name
			client_name = trim(client_name)
			EMReadScreen next_revw_dt, 8, rept_row, 42						'grabbing next review date
			next_revw_dt = replace(next_revw_dt, " ", "/")
			objExcel.Cells(excel_row, 2).Value = client_name				'adding read variables to the spreadsheet
			objExcel.Cells(excel_row, 3).Value = next_revw_dt				'adding read variables to the spreadsheet
			excel_row = excel_row + 1
		END IF
		rept_row = rept_row + 1
	LOOP UNTIL rept_row = 19								'looping until the script reads through the bottom of the page.
	PF8														'pf8 navigates to next page of ACTV
	rept_row = 7											'resetting the row to the top of the page.
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
LOOP UNTIL last_page = "THIS IS THE LAST PAGE"

excel_row = 2												'resetting excel row so script can review each case number found in previous loops.
DO
	back_to_SELF
	MAXIS_case_number = objExcel.Cells(excel_row, 1).Value					'reading case number from excel spreadsheet
	CALL find_variable("Environment: ", production_or_inquiry, 10)			'reading if script was started in production of inquiry, this is used later to navigate back from MMIS.
	CALL navigate_to_MAXIS_screen("ELIG", "HC")
	hhmm_row = 8															'setting starting point to review all HH members in ELIG HC
	DO																		'the script will now navigate to ELIG HC and begin to search for MA caes with DP as the elig type.
		EMReadScreen hc_type, 2, hhmm_row, 28
		IF hc_type = "MA" THEN												'if it finds MA as the HC type it will go into those results
			EMWriteScreen "X", hhmm_row, 26
			transmit
			EMReadScreen elig_type, 2, 12, 72
			IF elig_type = "DP" THEN										'once in those HC results it will look for DP as the elig type. DP is for MA-EPD
				EMWriteScreen "X", 9, 76
				transmit
				EMReadScreen pct_fpg, 4, 18, 38								'here it will check the percert of FPG client is at.
				pct_fpg = trim(pct_fpg)
				pct_fpg = pct_fpg * 1
				IF pct_fpg < 201 THEN										'If the client is 200% or under they may eligible for reimbursement
					PF3														'the script will now grab that person's member number and head into memb to get that person's PMI this will be used later to check MMIS
					PF3
					EMReadScreen hh_memb_num, 2, hhmm_row, 3
					CALL navigate_to_MAXIS_screen("STAT", "MEMB")
					EMWriteScreen hh_memb_num, 20, 76
					transmit
					EMReadScreen cl_pmi, 8, 4, 46
					cl_pmi = replace(cl_pmi, " ", "")
					DO
						IF len(cl_pmi) <> 8 THEN cl_pmi = "0" & cl_pmi
					LOOP UNTIL len(cl_pmi) = 8
					navigate_to_MMIS										'the script will now take the PMI and go into MMIS and check RELG
					DO
						EMReadScreen RKEY, 4, 1, 52
						IF RKEY <> "RKEY" THEN EMWaitReady 0, 0
					LOOP UNTIL RKEY = "RKEY"
					EMWriteScreen "I", 2, 19
					EMWriteScreen cl_pmi, 4, 19
					transmit
					EMWriteScreen "RELG", 1, 8
					transmit

					'Reading RELG to determine if the CL is active on MA-EPD
					EMReadScreen prog01_type, 8, 6, 13
						EMReadScreen elig01_type, 2, 6, 33
						EMReadScreen elig01_end, 8, 7, 36
					EMReadScreen prog02_type, 8, 10, 13
						EMReadScreen elig02_type, 2, 10, 33
						EMReadScreen elig02_end, 8, 11, 36
					EMReadScreen prog03_type, 8, 14, 13
						EMReadScreen elig03_type, 2, 14, 33
						EMReadScreen elig03_end, 8, 15, 36
					EMReadScreen prog04_type, 8, 18, 13
						EMReadScreen elig04_type, 2, 18, 33
						EMReadScreen elig04_end, 8, 19, 36

					IF ((prog01_type = "MEDICAID" AND elig01_type = "DP" AND elig01_end = "99/99/99") OR _
						(prog02_type = "MEDICAID" AND elig02_type = "DP" AND elig02_end = "99/99/99") OR _
						(prog03_type = "MEDICAID" AND elig03_type = "DP" AND elig03_end = "99/99/99") OR _
						(prog04_type = "MEDICAID" AND elig04_type = "DP" AND elig04_end = "99/99/99")) THEN

						EMWriteScreen "RMCR", 1, 8							'the script will now check RMCR for an active medicare case
						transmit

						'-----CHECKING FOR ON-GOING MEDICARE PART B-----
						EMReadScreen part_b_begin01, 8, 13, 4
							part_b_begin01 = trim(part_b_begin01)
						EMReadScreen part_b_end01, 8, 13, 15
						EMReadScreen part_b_begin02, 8, 14, 4
							part_b_begin02 = trim(part_b_begin02)
						EMReadScreen part_b_end02, 8, 14, 15

						IF (part_b_begin01 <> "" AND part_b_end01 = "99/99/99") THEN				'lastly the script will check RBYB to see what the client's buy in status is
							EMWriteScreen "RBYB", 1, 8
							transmit

							EMReadScreen accrete_date, 8, 5, 66
							EMReadScreen delete_date, 8, 6, 65
							accrete_date = replace(accrete_date, " ", "")

							IF ((accrete_date = "") OR (accrete_date <> "" AND delete_date <> "99/99/99")) THEN				'if the PMI is found to be open on MA-EPD, under 200% open on medicare and they don't have an end date on the delete date (rbyb) the script marks them as eligible for reimbursement.
								objExcel.Cells(excel_row, 4).Value = objExcel.Cells(excel_row, 4).Value & ("MEMB " & hh_memb_num & " ELIG FOR REIMBURSEMENT, ")  'writing eligibility status in spreadsheet
							END IF
							CALL write_value_and_transmit("RKEY", 1, 8)
						END IF
					ELSE
						CALL write_value_and_transmit("RKEY", 1, 8)
					END IF
					CALL navigate_to_MAXIS(production_or_inquiry)				'the script now navigates back to the environment the user left MAXIS in to continue searching Household members on the current case.
					hhmm_row = hhmm_row + 1
					CALL navigate_to_MAXIS_screen("ELIG", "HC")
				ELSE
					DO
						EMReadScreen at_hhmm, 4, 3, 51						'making sure the script made it back to ELIG/HC
						IF at_hhmm <> "HHMM" THEN PF3
					LOOP UNTIL at_hhmm = "HHMM"
					hhmm_row = hhmm_row + 1									'adding to the read row since we have finished evaluating this particular HH member.
				END IF
			ELSE
				PF3															'if the MA elig results don't have DP we end up here
				hhmm_row = hhmm_row + 1										'adding to the read row since we have finished evaluating this particular HH member.
			END IF
		ELSE
			hhmm_row = hhmm_row + 1											'If the elig/hc results aren't MA we end up here and add to the read row since we have finished evaluating this particular HH member.
		END IF
		IF hhmm_row = 20 THEN												'here we are determining that we've read all of the HH members on the current HHMM screen.
			PF8																'pf8 will cause elig hc to move to the next set of HH members if that page is full
			EMReadScreen this_is_the_last_page, 21, 24, 2					'if the script has read everyone on a page and PF8'd and reached the last page the script is done evaulating this case
		END IF
	LOOP UNTIL hc_type = "  " OR this_is_the_last_page = "THIS IS THE LAST PAGE"
	'Deleting the blank results to clean up the spreadsheet
	IF objExcel.Cells(excel_row, 4).Value = "" THEN
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete
		excel_row = excel_row - 1
	END IF
	excel_row = excel_row + 1										'the script adds 1 to the excel row to move onto the next case to evaluate
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""

FOR i = 1 to 4							'making the columns stretch to fit the widest cell
	objExcel.Columns(i).AutoFit()
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!!")
