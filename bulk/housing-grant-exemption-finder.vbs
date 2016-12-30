'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - HOUSING GRANT EXEMPTION FINDER.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "526"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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

'DIALOGS----------------------------------------------------------------------
BeginDialog Housing_grant_exemption_finder_dialog, 0, 0, 218, 120, "Housing Grant Exemption Finder"
  EditBox 84, 20, 130, 15, worker_number
  CheckBox 4, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 109, 100, 50, 15
    CancelButton 164, 100, 50, 15
  Text 4, 25, 65, 10, "Worker(s) to check:"
  Text 4, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 14, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 4, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Shows dialog
Do
	Do
		Dialog Housing_grant_exemption_finder_dialog
		If buttonpressed = cancel then stopscript
		If (all_workers_check = 0 AND worker_number = "") then MsgBox "Please enter at least one worker number." 	'allows user to select the all workers check, and not have worker number be ""
	LOOP until all_workers_check = 1 or worker_number <> ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Fun with dates! --Creating variables for the rolling 12 calendar months
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month -2'
CM_minus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
CM_minus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)
'current month -3'
CM_minus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", -3, date)            ), 2)
CM_minus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", -3, date)            ), 2)
'current month -4'
CM_minus_4_mo =  right("0" &             DatePart("m",           DateAdd("m", -4, date)            ), 2)
CM_minus_4_yr =  right(                  DatePart("yyyy",        DateAdd("m", -4, date)            ), 2)
'current month -5'
CM_minus_5_mo =  right("0" &             DatePart("m",           DateAdd("m", -5, date)            ), 2)
CM_minus_5_yr =  right(                  DatePart("yyyy",        DateAdd("m", -5, date)            ), 2)
'current month -6'
CM_minus_6_mo =  right("0" &             DatePart("m",           DateAdd("m", -6, date)            ), 2)
CM_minus_6_yr =  right(                  DatePart("yyyy",        DateAdd("m", -6, date)            ), 2)
'current month -7'
CM_minus_7_mo =  right("0" &             DatePart("m",           DateAdd("m", -7, date)            ), 2)
CM_minus_7_yr =  right(                  DatePart("yyyy",        DateAdd("m", -7, date)            ), 2)
'current month -8'
CM_minus_8_mo =  right("0" &             DatePart("m",           DateAdd("m", -8, date)            ), 2)
CM_minus_8_yr =  right(                  DatePart("yyyy",        DateAdd("m", -8, date)            ), 2)
'current month -9'
CM_minus_9_mo =  right("0" &             DatePart("m",           DateAdd("m", -9, date)            ), 2)
CM_minus_9_yr =  right(                  DatePart("yyyy",        DateAdd("m", -9, date)            ), 2)
'current month -10'
CM_minus_10_mo =  right("0" &            DatePart("m",           DateAdd("m", -10, date)           ), 2)
CM_minus_10_yr =  right(                 DatePart("yyyy",        DateAdd("m", -10, date)           ), 2)
'current month -11'
CM_minus_11_mo =  right("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'Establishing value of variables for the rolling 12 months
current_month = CM_mo & "/" & CM_yr
current_month_minus_one = CM_minus_1_mo & "/" & CM_minus_1_yr
current_month_minus_two = CM_minus_2_mo & "/" & CM_minus_2_yr
current_month_minus_three = CM_minus_3_mo & "/" & CM_minus_3_yr
current_month_minus_four = CM_minus_4_mo & "/" & CM_minus_4_yr
current_month_minus_five = CM_minus_5_mo & "/" & CM_minus_5_yr
current_month_minus_six = CM_minus_6_mo & "/" & CM_minus_6_yr
current_month_minus_seven = CM_minus_7_mo & "/" & CM_minus_7_yr
current_month_minus_eight = CM_minus_8_mo & "/" & CM_minus_8_yr
current_month_minus_nine = CM_minus_9_mo & "/" & CM_minus_9_yr
current_month_minus_ten = CM_minus_10_mo & "/" & CM_minus_10_yr
current_month_minus_eleven = CM_minus_11_mo & "/" & CM_minus_11_yr

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "REF #"
ObjExcel.Cells(1, 5).Value = "EMPS"
ObjExcel.Cells(1, 6).Value = "DISA DATES"
ObjExcel.Cells(1, 7).Value = "MFIP BEGIN DATE"
ObjExcel.Cells(1, 8).Value = current_month					'using date calculations above, list will generate a rolling 12 months of issuances
ObjExcel.Cells(1, 9).Value = current_month_minus_one
ObjExcel.Cells(1, 10).Value = current_month_minus_two
ObjExcel.Cells(1, 11).Value = current_month_minus_three
ObjExcel.Cells(1, 12).Value = current_month_minus_four
ObjExcel.Cells(1, 13).Value = current_month_minus_five
ObjExcel.Cells(1, 14).Value = current_month_minus_six
ObjExcel.Cells(1, 15).Value = current_month_minus_seven
ObjExcel.Cells(1, 16).Value = current_month_minus_eight
ObjExcel.Cells(1, 17).Value = current_month_minus_nine
ObjExcel.Cells(1, 18).Value = current_month_minus_ten
ObjExcel.Cells(1, 19).Value = current_month_minus_eleven
objExcel.cells(1, 20).Value = "Privileged Cases"

FOR i = 1 to 20		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT


'Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 21 'Starting with 21 because cols 1-20 are already used

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'establishing the row to start searching in the Excel spreadsheet
excel_row = 2

For each worker in worker_array
	back_to_self
	EMWriteScreen CM_mo, 20, 43				'
	EMWriteScreen CM_yr, 20, 46
	Call navigate_to_MAXIS_screen("REPT", "MFCM")			'navigates to MFCM in the current footer month/year'
	EMWriteScreen worker, 21, 13
	transmit

	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason----'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
    has_content_check = trim(has_content_check)
	If has_content_check <> "" then
		Do
			MAXIS_row = 7	'Sets the row to start searching in MAXIS for
			Do
				EMReadScreen emps_status, 2, MAXIS_row, 52		'Reading Emps Status & only searches for exempt emps status codes
				If  emps_status = "02" OR emps_status = "07" OR _
					emps_status = "08" OR emps_status = "12" OR _
					emps_status = "23" OR emps_status = "24" OR _
					emps_status = "27" OR emps_status = "15" OR _
					emps_status = "18" OR emps_status = "30" OR _
					emps_status = "33" THEN
						EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6  	'Reading case number
						EMReadScreen emps_status, 2, MAXIS_row, 52	'Reading emps_status
						'if more than one HH member is on the list then non-MEMB 01's don't have a case number listed, this fixes that
						If trim(MAXIS_case_number) = "" AND trim(client_name) <> "" then 			'if there's a name and no case number
							EMReadScreen alt_case_number, 8, MAXIS_row - 1, 6				'then it reads the row above
							MAXIS_case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'
						END IF

						'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
						If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
						all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
						If trim(MAXIS_case_number) = "" and trim(client_name) = "" then exit do			'Exits do if we reach the end

					'add case/case information to Excel
        			ObjExcel.Cells(excel_row, 1).Value = worker
        			ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
    				ObjExcel.Cells(excel_row, 5).Value = emps_status
					excel_row = excel_row + 1	'moving excel row to next row'
					'Blanking out variable
					MAXIS_case_number = ""
				END IF
				MAXIS_row = MAXIS_row + 1	'adding one row to search for in MAXIS
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Now the script goes back into MFCM and grabs the member # and client name, then cchecks the potentially exempt members for subsidized housing
excel_row = 2           're-establishing the row to start checking the members for
Do
	MAXIS_case_number = objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
	Call navigate_to_MAXIS_screen("REPT", "MFCM")
	EMWriteScreen "________", 20, 28					'clears case number
	EMWriteScreen MAXIS_case_number, 20, 28					'enters case number
	EMWriteScreen "x", 7, 36		'going into the SANC panel to get case info
	transmit
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete				'row gets deleted since it will get added to the priv case list at end of script in col 20
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
	END IF
	'For all of the cases that aren't privileged...
	EMReadScreen client_name, 50, 4, 16		'Reading client name
	client_name = trim(client_name)
	EMReadScreen memb_number, 2, 4, 12		'reading member number
	PF3

	If len(memb_number) = 1 then memb_number = "0" & right(memb_number, 2)		'adds 0 to member number and reads the right 2 digits
	client_name = trim(client_name)							'trims client name
	If MAXIS_case_number = "" then exit do						'exits do if the case number is ""

	ObjExcel.Cells(excel_row, 3).Value = client_name		'adds client name to Excel list
    ObjExcel.Cells(excel_row, 4).Value = memb_number		'adds client member number to Excel list

    'checking member for subsidized housing
    Call navigate_to_MAXIS_screen("STAT", "SHEL")
    EMWriteScreen memb_number, 20, 76						'enters member number as this is a person based panel
    transmit
    EmReadScreen sub_housing, 1, 6, 46
    If sub_housing <> "Y" then 				'if member doesn't have sub housing, then the sub housing fiater is not necessary
        'Deleting the blank results to clean up the spreadsheet
        SET objRange = objExcel.Cells(excel_row, 1).EntireRow
        objRange.Delete
        excel_row = excel_row - 1		'does not advance one row if the case gets deleted
    END IF
    excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

'Now the script checks for MFIP start date, disa dates
excel_row = 2           're-establishing the row to start checking the members for
DO
	MAXIS_case_number = objExcel.cells(excel_row, 2).Value	're-establishing the case numbers
	If MAXIS_case_number = "" then exit do					'if case number is blank then exits do loop
	Call navigate_to_MAXIS_screen("STAT", "PROG")
	'reading the MFIP start date
	EMReadScreen prog_one, 2, 6, 67				'checking 1st line of CASH PROG for elig MFIP
	EMReadScreen prog_status_one, 4, 6, 74
	EMReadScreen elig_begin_date_one, 8, 6, 44
	If prog_one = "MF" and prog_status_one = "ACTV" then
		elig_begin_date_one = Replace(elig_begin_date_one," ","/")
		elig_begin_date = elig_begin_date_one
	ELSE
		EMReadScreen prog_two, 2, 7, 67				'checking 2nd line of CASH PROG for elig MFIP if 1st line isn't MFIP
		EMReadScreen prog_status_two, 4, 7, 74
		EMReadScreen elig_begin_date_two, 8, 7, 44
		elig_begin_date = elig_begin_date_two
	END IF
	'enters elig begin dates into Excel
	ObjExcel.Cells(excel_row, 7).Value = elig_begin_date

	'Because some cases don't have HCRE dates listed, so when you try to go past PROG the script gets caught up. Do...loop handles this instance.
	PF3		'exits PROG to prompt HCRE if HCRE isn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"		'repeats until case is not in the HCRE panel

	'STAT DISA PORTION
	Call navigate_to_MAXIS_screen("STAT", "DISA")
	EMWriteScreen memb_number, 20, 76				'enters member number
	transmit
	'Reading the disa dates
	EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
	EMReadScreen disa_end_date, 10, 6, 69
	disa_start_date = Replace(disa_start_date," ","/")		'cleans up DISA dates
	disa_end_date = Replace(disa_end_date," ","/")
	disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
	If disa_dates = "__/__/____ - __/__/____" then disa_dates = "NO DISA INFO"
	'adding disa date to Excel
	ObjExcel.Cells(excel_row, 6).Value = disa_dates

	excel_row = excel_row + 1	'moving excel row to next row'
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list is complete

'Now the script inputs the payment information from INQD/INQX
excel_row = 2           're-establishing the row to start checking issuances for
DO
	MAXIS_case_number = objExcel.cells(excel_row, 2).Value	're-establishing the case numbers
	If MAXIS_case_number = "" then exit do
	back_to_self
	EMWriteScreen "________", 18, 43				'blanking out case number
	EMWriteScreen MAXIS_case_number, 18, 43				'adding case number
	Call navigate_to_MAXIS_screen("MONY", "INQX")
	EMWriteScreen CM_minus_11_mo, 6, 38		'entering footer month/year 11 months prior to see full year
	EMWriteScreen CM_minus_11_yr, 6, 41
	EMWriteScreen CM_mo, 6, 53		'entering current footer month/year
	EMWriteScreen CM_yr, 6, 56
	EMWriteScreen "x", 10, 5		'selecting MFIP
	transmit

	'creating an array of issuance months to fill in the Excel list from INQD
	issuance_months_array = array(current_month, current_month_minus_one, current_month_minus_two, current_month_minus_three, _
	current_month_minus_four, current_month_minus_five, current_month_minus_six, current_month_minus_seven, current_month_minus_eight, _
	current_month_minus_nine, current_month_minus_ten, current_month_minus_eleven)

	'searching for the housing grant issued on the INQX/INQD screen(s)
	excel_col = 8		'establishing the col to start at

	For each issuance_month in issuance_months_array 		'For next searches issuances for all rolling 12 months
		DO
			If issuance_month = "05/15" or issuance_month = "06/15" then 		'enters N/A for months prior to 07/15 since housing grant began in 07/15
				month_total = "n/a"
				exit do
			END IF
			row = 6				'establishing the row to start searching for issuance'
			DO
				EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
				If housing_grant = "  " then exit do		'exits the do loop once the end of the issuances is reached
				IF housing_grant = "HG" then
					'reading the housing grant information
					EMReadScreen HG_amt_issued, 7, row, 40
					EMReadScreen HG_month, 2, row, 73
					EMReadScreen HG_year, 2, row, 79
					INQD_issuance_month = HG_month & "/" & HG_year		'creates a new varible for HG month and year
					If issuance_month = INQD_issuance_month then 		'if the issuance found matches the issuance month then
						HG_amt_issued = trim(HG_amt_issued)				'trims the HG amt issued variable
					ELSE
						HG_amt_issued = "0"
					END IF
					If month_total = "" then month_total = "0" 			'establishes a value for a "" amount found
					month_total = month_total + abs(HG_amt_issued)		'adds HG amt issued to variable incrementally
				END IF
				row = row + 1											'adds one row to search the next row
			Loop until row = 18											'repeats until the end of the page
			PF8
			EMReadScreen last_page_check, 21, 24, 2
			If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
		LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

		'adding issuance amounts to the appropriate column
		ObjExcel.Cells(excel_row, excel_col).Value = month_total			'adds month_total to the Excel
		excel_col = excel_col + 1
		month_total = ""	'resets the month_total variable to "" so that values don't get pulled into next loop

		'this do...loop gets the user back to the 1st page on the INQD screen to check the next issuance_month
		Do
			PF7
			EMReadScreen first_page_check, 20, 24, 2
		LOOP until first_page_check = "THIS IS THE 1ST PAGE"	'keeps hitting PF7 until user is back at the 1st page
	NEXT

	excel_row = excel_row + 1	'moving excel row to next row
	STATS_counter = STATS_counter + 1		 'adds one instance to the stats counter
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list is complete

'Creating the list of privileged cases and adding to the spreadsheet
prived_case_array = split(priv_case_list, "|")
excel_row = 2				'establishes the row to start writing the PRIV cases to

FOR EACH MAXIS_case_number in prived_case_array
	objExcel.cells(excel_row, 20).value = MAXIS_case_number		'inputs cases into Excel
	excel_row = excel_row + 1								'increases the row
NEXT

'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

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

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the list generated.")
