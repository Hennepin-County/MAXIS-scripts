'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - AVS Panel Report.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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

call changelog_update("04/25/2025", "Initial version.", "Dave Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Constants
worker_col  = 1
case_col   = 2
name_col   = 3
memb_col   = 4
pmi_col    = 5
program_col = 6
review_col = 7
status_col = 8
elig_col = 9

'Connects to BlueZone
EMConnect ""
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 286, 130, "Pull REPT data into Excel dialog"
  EditBox 135, 20, 145, 15, worker_number
  CheckBox 70, 60, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 70, 70, 190, 10, "Check here to resume from a previous spreadsheet.", resume_check
  ButtonGroup ButtonPressed
    OkButton 185, 110, 45, 15
    CancelButton 235, 110, 45, 15
  Text 70, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 70, 85, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 110, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 25, 65, 10, "Worker(s) to check:"
EndDialog

Do
	Do
		Dialog dialog1
		cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then MsgBox "Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
	LOOP until all_workers_check = 1 or worker_number <> ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If resume_check = 1 Then
	DO
		call file_selection_system_dialog(excel_file_path, ".xlsx")

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
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 186, 65, "Dialog"
	ButtonGroup ButtonPressed
	  OkButton 75, 40, 50, 15
	  CancelButton 130, 40, 50, 15
	EditBox 110, 15, 70, 15, excel_row
	Text 10, 20, 90, 10, "Excel row to start from:"
  	EndDialog
	Do
	    Do
			Dialog dialog1
			cancel_without_confirmation
			If isnumeric(excel_row) = false then MsgBox "Please enter the excel row to begin checking from." 'allows user to select the all workers check, and not have worker number be ""
	    LOOP until isnumeric(excel_row) = true and excel_row > 1
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End If 
'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)
If resume_check <> checked Then 
	'Opening the Excel file
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = True

	'Setting the first 4 col as worker, case number, name, and APPL date
	ObjExcel.Cells(1, 1).Value = "WORKER"
	ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
	ObjExcel.Cells(1, 3).Value = "NAME"
	ObjExcel.Cells(1, 4).Value = "MEMBER NUMBER"
	ObjExcel.Cells(1, 5).Value = "PMI"
	ObjExcel.Cells(1, 6).Value = "MAJOR PROGRAM"
	ObjExcel.Cells(1, 7).Value = "REVIEW DATE"
	ObjExcel.Cells(1, 8).Value = "AVSA STATUS"
	ObjExcel.Cells(1, 9).Value = "ELIGIBILITY TYPE"
	FOR i = 1 to 9	'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
	NEXT
	ObjExcel.ActiveWorkbook.SaveAs "C:\MAXIS-scripts\AVS Panel Report.xlsx"
	'Constants


	'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
	If all_workers_check = checked then
		call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
	Else		'If worker numbers are litsted - this will create an array of workers to check
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
	End if

	'Setting the variable for what's to come
	excel_row = 2
	all_case_numbers_array = "*"

	For each worker in worker_array
		worker = trim(ucase(worker))					'Formatting the worker so there are no errors
		back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
		'Exclude caseloads that shouldn't have assets to save time
		If worker <> "X127EN6" and worker <> "X127EZ5" and worker <> "X127FG2" and worker <> "X127EW4" and worker <> "X1274EC" and worker <> "X127FG1" and worker <> "X127F3K" and worker <> "X127F3P" and worker <> "X127F3F" and worker <> "X127F4E" and worker <> "X127CCL" Then
			Call navigate_to_MAXIS_screen("rept", "actv")
			EMWriteScreen worker, 21, 13
			TRANSMIT
			EMReadScreen user_worker, 7, 21, 71		'
			EMReadScreen p_worker, 7, 21, 13
			IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

			'msgbox "worker " & worker

			IF worker_number = "X127CCL" or worker = "127CCL" THEN
				DO
					EmReadScreen worker_confirmation, 20, 3, 11 'looking for CENTURY PLAZA CLOSED
					EMWaitReady 0, 0
					'MsgBox "Are we waiting?"
				LOOP UNTIL worker_confirmation = "CENTURY PLAZA CLOSED"
			END IF

			'Skips workers with no info
			EMReadScreen has_content_check, 1, 7, 8
			If has_content_check <> " " then
				PF5 ' Sort by case number so we don't lose place going in and out of HC popup
				'Grabbing each case number on screen
				Do
				    'Set variable for next do...loop
					MAXIS_row = 7
					'Checking for the last page of cases.
					EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
					EMReadscreen number_of_pages, 4, 3, 76 'getting page number because to ensure it doesnt fail'
					number_of_pages = trim(number_of_pages)
					Do
						EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
						'EMReadScreen client_name, 21, MAXIS_row, 21			'Reading client name
						'EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
						EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status

						'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
						MAXIS_case_number = trim(MAXIS_case_number)
						If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
						all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

						If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

						'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
						If HC_status <> " " and HC_status <> "I" and HC_status <> "P" then 
							add_case_info_to_Excel = True
							Call write_value_and_transmit("X", MAXIS_row, 3)	'Writing to the HC status field to get the status
							EMReadScreen popup_check, 6, 2, 41 'Make sure we aren't still on the REPT screen
							If popup_check = "Active" Then
								For memb_row = 7 to 20
									EMReadScreen Memb_number, 2, memb_row, 5	'Reading member number
									If memb_number <> "  " then
										EMReadScreen memb_pmi, 7, memb_row, 9
										EMReadScreen memb_name, 25, memb_row, 18
										EMReadScreen memb_programs, 20, memb_row, 50	'Reading programs
										If trim(memb_programs) <>"IMD" then 
											'Write the info to excel
											ObjExcel.Cells(excel_row, 1).Value = worker
											ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
											ObjExcel.Cells(excel_row, 3).Value = trim(memb_name)
											ObjExcel.Cells(excel_row, 4).Value = trim(memb_number)
											ObjExcel.Cells(excel_row, 5).Value = trim(memb_pmi)
											ObjExcel.Cells(excel_row, 6).Value = trim(memb_programs)
											'IF next_revw_date <> "        " THEN ObjExcel.Cells(excel_row, 7).Value = replace(next_revw_date, " ", "/")
											excel_row = excel_row + 1
										End If
									End if
								next
								PF3 'Pop back to the REPT/ACTV screen 
								MAXIS_row = 7	'Resetting the row to the first one
							End if 
						End If 

						MAXIS_row = MAXIS_row + 1
						add_case_info_to_Excel = ""	'Blanking out variable
						MAXIS_case_number = ""			'Blanking out variable
						STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					Loop until MAXIS_row = 19
					PF8
				Loop until last_page_check = "THIS IS THE LAST PAGE"
			END IF
		End If

	next 
	ObjExcel.ActiveWorkbook.Save
	excel_row = 2 ''Resetting the excel row to 2, as we are going to be checking the cases now
End If 
'If collect_COLA_stats = True then 'Replace this with logic to choose existing file, resume, etc.


save_count = 0	'Resetting the save count
'This loop will navigate to UNEA and check each case for the specified types of income
Do
	'Assign case number from Excel
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2)
	memb_number = ObjExcel.Cells(excel_row, 4)	'Reading member number from the Excel sheet
	if len(memb_number) = 1 then memb_number = "0" & memb_number 'make sure member number is 2 digits for maxis
	'Exiting if the case number is blank
	If MAXIS_case_number = "" then exit do
	'Navigate to STAT/UNEA for said case number
	call navigate_to_MAXIS_screen("STAT", "MEMB")
	Call write_value_and_transmit(memb_number, 20, 76)	'Writing member number
	EMReadScreen member_age, 2, 8, 76		'Checking if the member is under 18
	If member_age = "  " then member_age = 0	'If the age is blank, set it to 0 so we can do math on it
	If member_age = 17 Then	EMReadScreen birth_month, 2, 8, 42 
	If member_age > 16 Then 	
		Call navigate_to_MAXIS_screen("STAT", "REVW")
		EMReadScreen review_date, 8, 9, 70
		review_date = replace(review_date, " ", "/")	'Formatting the date to be more readable
		ObjExcel.Cells(excel_row, review_col).Value = review_date
		If member_age = 17 Then	'If the member is 17, we need to check the birth month
			If birth_month <= left(review_date, 2) then member_age = 18	'If the birth month is less than or equal to the review date, they are 18 for avsa purposes
		End If
	End If 
	If member_age >= 18 Then 'This will be handling for people 18 or over at review, that need AVSA panels
		Call navigate_to_MAXIS_screen("STAT", "MEMI")
		EMReadScreen marital_status, 1, 7, 40	'Reading marital status
		Call navigate_to_MAXIS_screen("STAT", "AVSA")
		call write_value_and_transmit(memb_number, 20, 76)	'Writing member number
		EMReadScreen panel_status, 1, 2, 78	'Reading panel status
		If panel_status = "0" Then
			'If no panel, we need to check elig type to see if they have asset test
			Call navigate_to_MAXIS_screen("ELIG", "HC")
			EMReadScreen warning_msg, 50, 24, 2
			If Instr(warning_msg, "INVALID FOR PERIOD") <> 0 Then
				EMWriteScreen CM_mo, 20, 43
				EMWriteScreen CM_yr, 20, 46
				transmit
			End If
			For hc_row = 8 to 19
				EMReadScreen ref_numb, 2, hc_row, 3	'Reading reference number 
				If ref_numb = memb_number Then 
					EMReadScreen prog_status, 3, hc_row, 68
					If prog_status <> "APP" Then                        'Finding the approved version
						EMReadScreen total_versions, 2, hc_row, 64
						If total_versions = "01" or total_versions = "  " Then
							footer_to_use = date
							For fm = 0 to -7
								footer_to_use = dateadd("m", fm, footer_to_use)	'Adding months to the date
								footer_month = datepart("m", footer_to_use)	'Getting the month
								footer_year = datepart("yyyy", footer_to_use)	'Getting the year
								footer_month = right("0" & footer_month, 2)	'Formatting the month to be two digits
								footer_year = right(footer_year, 2)			'Formatting the year to be two digits
								EMWriteScreen footer_month, 20, 56
								EMWriteScreen footer_year,	20, 59
								Transmit
								EMReadScreen prog_status, 3, hc_row, 68
								If prog_status = "APP" Then
									exit for
								Else 
									EMReadScreen total_versions, 2, hc_row, 64
									If total_versions <> "01" Then
										For current_version = right(total_versions, 1) to 1 step -1
											Call write_value_and_transmit(current_version, hc_row, 59)	'Writing the current version to the screen
											EmReadScreen prog_status, 3, hc_row, 68
											If prog_status = "APP" Then exit for
										Next
									End IF 		
								End If 	
								If prog_status = "APP" Then exit for 'If we find the approved version, we can exit the loop						
							Next 
						End if 
					End if 
					'We should be on the approved version now, so we can check for the eligibility type
					approved_progs_string = ""	'Blanking out the approved programs string
					EMReadScreen major_prog, 4, hc_row, 28	'Reading eligibility type
					major_prog = trim(major_prog)
					If major_prog = "MA" or major_prog = "EMA" Then
						Call write_value_and_transmit("X", hc_row, 26)
						EMReadScreen elig_type, 2, 11, 72	'Reading eligibility type
						approved_progs_string = approved_progs_string & " " & major_prog & "/" & Elig_type	'Adding the approved program to the string
						PF3
					End If 
					For i = 1 to 5 'checking for additional progs with this loop
						EMReadScreen check_progs, 2, hc_row + i, 3							
						If check_progs <> "  " Then
							 Exit For	'Exit loop at next member
						Else 
							EmReadScreen prog_status, 3, hc_row + i, 68
							If prog_status = "APP" Then 
								EMReadScreen major_prog, 4, hc_row + i, 28	'Reading eligibility type
								If trim(major_prog) <> "MA" Then approved_progs_string = approved_progs_string & " " & trim(major_prog)	'Adding the approved program to the string
							Else
								EMReadScreen total_versions, 2, hc_row, 64
								If total_versions <> "01" and total_versions <> "  " Then
									For current_version = right(total_versions, 1) to 1 step -1
										Call write_value_and_transmit(current_version, hc_row, 59)	'Writing the current version to the screen
										EmReadScreen prog_status, 3, hc_row, 68
										If prog_status = "APP" Then 
											EMReadScreen major_prog, 4, hc_row + i, 28	'Reading eligibility type
											If trim(major_prog) <> "MA" Then approved_progs_string = approved_progs_string & " " & trim(major_prog)	'Adding the approved program to the string
											exit for
										End If 
									Next
								End IF 		
							End If 	
						End If
					NExt 	
				End if 
			Next 
			'Put the status on the excel sheet
		
			ObjExcel.Cells(excel_row, status_col).Value = "No AVSA panel"
			ObjExcel.Cells(excel_row, elig_col).Value = approved_progs_string	'Adding the approved programs to the excel sheet
		Else 'panel exists, read the status
			panel_status = ""	'Blanking out the variable
			For avsa_line = 9 to 14 'Read through lines of avsa, looking for requested or invalid status
				EMReadScreen role, 1, avsa_line, 62	'Reading role
				EMReadScreen status, 1, avsa_line, 76	'Reading status
				If status = "R" or status = "I" then panel_status = role & " " & status	& ", " 'If the status is "R", we need to add the role to the status
			Next
			If panel_status <> "" then
				panel_status = left(panel_status, len(panel_status) - 2)	'Removing the last comma and space from the string
				ObjExcel.Cells(excel_row, status_col).Value = "Invalid or missing forms: " & panel_status	'Adding the status to the excel sheet
			Else
				ObjExcel.Cells(excel_row, status_col).Value = "Valid AVSA panel"	'If the status is blank, we need to add a valid status to the excel sheet
			End If	
		End If 
	Else
		ObjExcel.Cells(excel_row, status_col).Value = "N/A"
	End If 
	save_count = save_count + 1	'Adding one to the save count, so we can save every 25 cases
	If save_count = 25 then
		ObjExcel.ActiveWorkBook.Save 'As "C:\MAXIS-scripts\AVS Panel Report.xlsx"
		save_count = 0	'Resetting the save count
	End If
	excel_row = excel_row + 1	'Advances to look at the next row
Loop until MAXIS_case_number = ""
'End if

'Autofitting columns
For col_to_autofit = 1 to 9
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure("Success! Your REPT/ACTV list has been created.")
