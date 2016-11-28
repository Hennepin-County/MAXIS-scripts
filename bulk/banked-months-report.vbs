'Hard coding that needs to be updated each year: MAXIS_footer_year, counted_date_year

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - BANKED MONTHS REPORT.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 219         'manual run time in seconds
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""		'connecting to MAXIS
Call get_county_code	'gets county name to input into the 1st col of the spreadsheet
developer_mode_checkbox = checked 	'defauting the person note option to NOT person note

report_date = MAXIS_footer_month & "/" & MAXIS_footer_year			'creating date variables to measure against person note counted dates

'Runs the dialog'
Do
	Do
		Do
			'The dialog is defined in the loop as it can change as buttons are pressed (populating the dropdown)'
			BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 406, 155, "Banked Month Report"
				EditBox 165, 50, 160, 15, banked_months_clients_excel_file_path
				ButtonGroup ButtonPressed
				  PushButton 330, 50, 45, 15, "Browse...", select_a_file_button
				CheckBox 155, 70, 200, 10, "Check here to run without Person Noting", developer_mode_checkbox
				DropListBox 215, 120, 140, 15, "select one..." & month_list, report_month_dropdown
				ButtonGroup ButtonPressed
				  OkButton 295, 135, 50, 15
				  CancelButton 350, 135, 50, 15
				Text 10, 10, 365, 10, "Select the Excel File that contains your list of clients that used banked months and need to be reported to DHS."
				Text 15, 25, 355, 15, "The file must be in the correct format for the script to operate. The template with the correct format can be found on Git Hub for download. Review the instructions on SIR for help with this."
				Text 10, 55, 150, 10, "Select an Excel file of banked months clients:"
				Text 10, 70, 140, 65, "Once the correct file is selected, the months will be listed to the right. You must select the month in which the banked months you are going to report were used. This will also select the footer month the script will look in for information. This month selected will also be used in creating the Excel File of the DHS report."
				Text 155, 110, 190, 10, "Select the month that you are creating the report for."
				Text 155, 125, 55, 10, "Month to Report:"
				Text 155, 85, 245, 20, "** Person Noting should only happen ONCE per report month in each County - leave checked unless you are sure Person Noting should happen."
			EndDialog
			err_msg = ""
			Dialog SNAP_Banked_Month_Report_Dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If banked_months_clients_excel_file_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
					month_list = ""	'Blanks the Month list out so that the previous worksheets are not still included'
				End If
				call file_selection_system_dialog(banked_months_clients_excel_file_path, ".xlsx") 'allows the user to select the file'
			End If
			If banked_months_clients_excel_file_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(banked_months_clients_excel_file_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		month_list = ""
		For Each objWorkSheet In objWorkbook.Worksheets
			month_list = month_list & chr(9) & objWorkSheet.Name
		Next
		If report_month_dropdown = "select one..." then err_msg = err_msg & vbNewLine & "You must select a month that you are running this script for."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'This vbOKCanel gives the user time to stop the script if they chose the incorrect person noting option. Ok to proceed, cancel to stop the script.
If developer_mode_checkbox = checked then
	person_noting = Msgbox("You have selected this script to NOT add a Person Note." & vbNewLine & "Note that this is the only way we have to track months a client has used a Banked Month." & vbNewLine & _
    "Check the instructions for further details on this option.", vbOkCancel + vbExclamation, "Person notes will NOT be added")
  	If person_noting = vbCancel then script_end_procedure("You have selected the cancel button, so the script has ended.")
Elseif developer_mode_checkbox = unchecked then
	person_noting = Msgbox("You have selected this script TO ADD a Person Note." & vbNewLine & "A Person Note WILL be added for EVERY Client added to the DHS Report.", vbOkCancel + vbExclamation, "Person notes WILL be added")
 	If person_noting = vbCancel then script_end_procedure("You have selected the cancel button, so the script has ended.")
END IF

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

objExcel.worksheets(report_month_dropdown).Activate			'Activates the selected worksheet'
report_month_dropdown = trim(report_month_dropdown)			'prevents matching errors so that there is not a miss in the footer month selection'

'This assigns a footer month and year based on the worksheet names selected in the dropdown from the dialog'
Select Case report_month_dropdown
Case "January 2016"
	MAXIS_footer_month = "01"
	MAXIS_footer_year = "16"
Case "February 2016"
	MAXIS_footer_month = "02"
	MAXIS_footer_year = "16"
Case "March 2016"
	MAXIS_footer_month = "03"
	MAXIS_footer_year = "16"
Case "April 2016"
	MAXIS_footer_month = "04"
	MAXIS_footer_year = "16"
Case "May 2016"
	MAXIS_footer_month = "05"
	MAXIS_footer_year = "16"
Case "June 2016"
	MAXIS_footer_month = "06"
	MAXIS_footer_year = "16"
Case "July 2016"
	MAXIS_footer_month = "07"
	MAXIS_footer_year = "16"
Case "August 2016"
	MAXIS_footer_month = "08"
	MAXIS_footer_year = "16"
Case "September 2016"
	MAXIS_footer_month = "09"
	MAXIS_footer_year = "16"
Case "October 2016"
	MAXIS_footer_month = "10"
	MAXIS_footer_year = "16"
Case "November 2016"
	MAXIS_footer_month = "11"
	MAXIS_footer_year = "16"
Case "December 2016"
	MAXIS_footer_month = "12"
	MAXIS_footer_year = "16"
End Select

'Sets up the array to store all the information for each client'
Dim Banked_Month_Client_Array ()
ReDim Banked_Month_Client_Array (14, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num          = 1			'Each of the case numbers will be stored at this position'
Const clt_pmi           = 2
Const memb_num          = 3
Const clt_name          = 4
Const clt_first_name    = 5
Const clt_last_name     = 6
Const comments          = 7
Const abawd_used        = 8
Const second_abawd_used = 9
Const abawd_count       = 10
Const second_count      = 11
Const send_to_DHS       = 12		'This is a True/False value that determines which report this array item will be entered on'
Const reason_excluded   = 13		'If the above is False, an explanation will be added to help the user check/track these error prone cases'
Const clt_filter        = 14

'Now the script adds all the clients on the excel list into an array
excel_row = 3 're-establishing the row to start checking the members for
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 4).Value          're-establishing the case numbers for functions to use
	If MAXIS_case_number = "" then exit do
	MAXIS_case_number = trim(MAXIS_case_number)
	client_first_name = objExcel.cells(excel_row, 3).Value		'Pulls the client's first and last names and trims for future matching
	client_last_name  = objExcel.cells(excel_row, 2).Value             're-establishing the client name
	client_first_name = UCase(trim(client_first_name))
	client_last_name  = UCase(trim(client_last_name))
	'Adding client information to the array'
	ReDim Preserve Banked_Month_Client_Array(14, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	Banked_Month_Client_Array (case_num,       entry_record) = MAXIS_case_number		'The client information is added to the array'
	Banked_Month_Client_Array (clt_last_name,  entry_record) = client_last_name
	Banked_Month_Client_Array (clt_first_name, entry_record) = client_first_name
	Banked_Month_Client_Array (clt_name,       entry_record) = client_first_name & " " & client_last_name
	Banked_Month_Client_Array (comments,       entry_record) = objExcel.cells(excel_row, 6).Value
	Banked_Month_Client_Array (send_to_DHS,    entry_record) = TRUE				'This is the default, this may be changed as info is checked'
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	excel_row = excel_row + 1
Loop

'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document and don't want the script to be confused
objExcel.Quit

'Now we will get PMI and Member Number for each client on the array.'
For item = 0 to UBound(Banked_Month_Client_Array, 2)
	MAXIS_case_number = Banked_Month_Client_Array(case_num,item)				'Case number is set for each loop as it is used in the FuncLib functions'
	Call navigate_to_MAXIS_screen("INFC", "WORK")						'Finding client information on STAT MEMB'
	EMReadScreen WORK_check, 4, 2, 51									'Making sure the script made it to INFC/WORK '
	IF WORK_check = "WORK" Then
		work_maxis_row = 7
		DO
			EMReadScreen client_referred, 26, work_maxis_row, 7			'Reads the client name from INFC/WORK'
			client_referred = trim(client_referred)
			IF Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name, item) = client_referred then
				memb_check = vbYes		'If the name on INFC/WORK exactly matches the name from the initial excel list, the script does not need user input and will gather the PMI and Reference Number'
				EMReadScreen Banked_Month_Client_Array(clt_pmi,  item), 8, work_maxis_row, 34
				EMReadScreen Banked_Month_Client_Array(memb_num, item), 2, work_maxis_row, 3
			ElseIf Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name, item) <> client_referred then 	'if name doesn't match the referral name the confirmation is required by the user
				memb_check = MsgBox ("Client listed on your report: " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name, item) & _
			  	vbNewLine &        "Client name listed in MAXIS: " & trim(client_referred) & vbNewLine & vbNewLine & "Is this the client you are reporting as using banked months?", vbYesNo + vbQuestion, "Confirm Client using Banked Monhts")
				If memb_check = vbYes Then		'If the user confirms that this is the correct client, the PMI and Ref number are gathered'
					EMReadScreen Banked_Month_Client_Array(clt_pmi,  item), 8, work_maxis_row, 34
					EMReadScreen Banked_Month_Client_Array(memb_num, item), 2, work_maxis_row, 3
				ElseIf memb_check = vbNo Then	'If the user says NO the script will see if there are other clients listed on INFC/WORK and start back at the beginning of the loop to try to match'
					EMReadScreen next_clt, 1, (work_maxis_row + 1), 7
					If next_clt = " " Then		'If no clients are matchs, the script removes this entry from the DHS report'
						MsgBox "There are no additional clients on this case that have had a workforce referral. Since banked months require E&T participation, there must be a referral This client - " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name, item) & _
				    	" - will not be added to the DHS report."		'The user is alerted to this'
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report'
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "Person not matched with name in MAXIS. | "		'Explanation for the rejected report'
						End If
					End If
				END IF
			work_maxis_row = work_maxis_row + 1		'Increments to read the next row for a new client'
			STATS_counter = STATS_counter + 1
		Loop until next_clt = " " OR memb_check = vbYes		'Loop is ended once there are no more clients on INFC/WORK OR a match has been made'
	Else																'If there is INFC/WORK for a client - there was no E&T referral done'
		Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
		Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "No Workforce1 referral was done. Banked Months requires client to participate in E&T, so a Workforce 1 Referral needs to be completed. | "
	End If
Next

'informational box for the users with next steps so they know what to expect
list_done_msgbox = MsgBox ("The script has finished compiling the list of clients to add to the Report." & vbNewLine & vbNewLine & _
  "It will now continue and do the following:" & vbNewLine & "* Check in STAT for some possible exemptions" & vbNewLine & _
  "* Get the list of Counted ABAWD Months (including 2nd set)" & vbNewLine & "* Add a person note that a banked month was counted" & vbNewLine & _
  "(Unless you checked for the script to NOT person note)" & vbNewLine & "* Add all clients that still appear to have used a Banked Month to the report" & vbNewLine & _
  "* Create a report of the clients that were NOT added to the report" & vbNewLine & vbNewLine & "The script will take a few minutes to check ELIG and STAT before asking you for the Excel File of the DHS Report", vbOkOnly + vbInformation, "Client List Created")

For item = 0 to UBound(Banked_Month_Client_Array, 2)		'Now each entry in the array will be checked in ELIG and STAT'
	MAXIS_case_number = Banked_Month_Client_Array(case_num,item)	'Case number is set for each loop as it is used in the FuncLib functions'
	If Banked_Month_Client_Array(send_to_DHS, item) = TRUE Then	'If a case has already been removed from the DHS report, no additional check is needed'
		call navigate_to_MAXIS_screen ("ELIG", "FS")		'Checking ELIG - this is footer month specific (set above)'
		EMReadScreen no_SNAP, 10, 24, 2
		If no_SNAP = "NO VERSION" then						'NO SNAP version means no banked months could have been used'
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
		END IF
		EMWriteScreen "99", 19, 78
		transmit
		'This brings up the FS versions of eligibilty results to search for approved versions
		status_row = 7
		Do
			EMReadScreen app_status, 8, status_row, 50
			If app_status = "        " then
				Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
				PF3
				exit do 	'if end of the list is reached then exits the do loop
			End if
			If app_status = "UNAPPROV" Then status_row = status_row + 1
		Loop until  app_status = "APPROVED" or app_status = "        "
			If app_status <> "APPROVED" then
				Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
			Elseif app_status = "APPROVED" then
				EMReadScreen vers_number, 1, status_row, 23
				EMWriteScreen vers_number, 18, 54
				transmit
				'now checking for banked months recipient elig on FSPR screen
				elig_maxis_row = 7
				Do
					EMReadScreen clt_on_snap, 2, elig_maxis_row, 10			'Each line of the elig results are checked to find the client that used banked months'
					IF clt_on_snap = Banked_Month_Client_Array(memb_num,item) Then
						EMReadScreen pers_elig, 8, elig_maxis_row, 57		'Once found, client eligibility is determined'
						If pers_elig = "ELIGIBLE" then exit do
						IF pers_elig = "INELIGIB" Then						'If ineligible the footer month, they did not use banked months'
							Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
							Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "Client listed as Ineligible for SNAP on ELIG/FS for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
							exit do
						End If
					ElseIf clt_on_snap = "  " Then							'If client is not found, they did not receive SNAP on this case in this month'
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "Client not listed on ELIG/FS for " &  MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
					Else
						elig_maxis_row = elig_maxis_row + 1
						If elig_maxis_row = 19 Then
							PF8
							elig_maxis_row = 7
						End If
					End If
				Loop until clt_on_snap = Banked_Month_Client_Array(memb_num, item) OR Banked_Month_Client_Array(send_to_DHS,item) = FALSE
				EMWriteScreen "FSB2", 19, 70
				transmit
				EMReadScreen fs_prorated, 8, 11,40			'Looking to see if benefits were prorated, if so, no banked months were used'
				IF fs_prorated = "Prorated" Then
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "SNAP is prorated in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
				Else
					'This section checks for cases that were not approved as PRORATED, but are prorated for the report date on the PROG panel
					Call navigate_to_MAXIS_screen("STAT", "PROG")
					EMReadScreen elig_month, 2, 10, 44
					EMReadScreen elig_date, 2, 10, 47
					EMReadScreen elig_year, 2, 10, 50
					prorated_date = elig_month & "/" & elig_year		'creating date variables to measure against report month
					If prorated_date = report_date then
						If elig_date <> "01" then
							Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
							Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "SNAP is prorated in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
						END if
					END IF
					'handling for cases that do not have a completed HCRE panel
	    			PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	    			Do
	    				EMReadscreen HCRE_panel_check, 4, 2, 50
						If HCRE_panel_check = "HCRE" then
	                    	PF10	'exists edit mode in cases where HCRE isn't complete for a member
	    					PF3
	    				END IF
	    			Loop until HCRE_panel_check <> "HCRE"
				END IF
			END If
		'///////SCRIPT WILL NOW CHECK FOR POSSIBLE EXPEMTIONS FOR CLIENT'
		'Age exemption'
		call navigate_to_MAXIS_screen ("STAT", "MEMB")																					'Cient age is listed on STAT MEMB'
		Call write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)			'Writes the clt reference number on the command line to get to the correct MEMB Panel'
		EMReadScreen cl_age, 2, 8, 76																													'Reads the client age'
		cl_age = abs(cl_age)																																	'Makes sure the age is seen in the script as a number for the math that come next'
		IF cl_age < 18 OR cl_age >= 50 THEN 																									'Compares to the age exclusions'
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE 	'Removing this client from DHS report - reason on next line'															'Codes the array to not include this client and the reason'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have exemption. Age = " & cl_age & ". | "
		End If
		'Exemptions for Disability'
		disa_status = false																																		'Resets the variable for the looping'
		call navigate_to_MAXIS_screen("STAT", "DISA")																							'Information about disability is on STAT/DISA'
		Call write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)				'Enters the clt reference number to get to the right disa panel'
		EMReadScreen num_of_DISA, 1, 2, 78
		IF num_of_DISA <> "0" THEN			'If there is a DISA panel, this code will check for an openended or future date disability or certification'
			EMReadScreen disa_end_dt, 10, 6, 69
			disa_end_dt = replace(disa_end_dt, " ", "/")
			EMReadScreen cert_end_dt, 10, 7, 69
			cert_end_dt = replace(cert_end_dt, " ", "/")
			IF IsDate(disa_end_dt) = True THEN
				IF DateDiff("D", date, disa_end_dt) > 0 THEN
					disa_status = True
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA end date = " & disa_end_dt & ". | "
				END IF
			ELSE
				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
					disa_status = True
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA has no end date. | "
				END IF
			END IF
			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
				IF DateDiff("D", date, cert_end_dt) > 0 THEN
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA Certification end date = " & cert_end_dt & ". | "
				End If
			ELSE
				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
					EMReadScreen cert_begin_dt, 8, 7, 47
					IF cert_begin_dt <> "__ __ __" THEN
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA certification has no end date. | "
					End If
				END IF
			END IF
		End If
		'Checking for earned income exemptions'
		'JOBS'
		prosp_inc = 0		'Variables are reset for each run of the loop'
		prosp_hrs = 0
		prospective_hours = 0

		CALL navigate_to_MAXIS_screen("STAT", "JOBS")		'Checking JOBS for income and hours'
		EmWriteScreen Banked_Month_Client_Array(memb_num, item), 20, 76
		EmWriteScreen "01", 20, 79
		transmit
		EMReadScreen num_of_JOBS, 1, 2, 78
		IF num_of_JOBS <> "0" THEN
			DO
				EMReadScreen jobs_end_dt, 8, 9, 49			'Looking to be sure job has not ended'
				EMReadScreen cont_end_dt, 8, 9, 73
				IF jobs_end_dt = "__ __ __" THEN
					CALL write_value_and_transmit("X", 19, 38)	'Information is gathered from the PIC'
					EMReadScreen prosp_monthly, 8, 18, 56
					prosp_monthly = trim(prosp_monthly)
					IF prosp_monthly = "" THEN prosp_monthly = 0	'Finds budeted income'
					prosp_inc = prosp_inc + prosp_monthly			'All budgeted income will be added together'
					EMReadScreen prosp_hrs, 8, 16, 50					'Looking for reported hours'
					IF prosp_hrs = "        " THEN prosp_hrs = 0
					prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
					EMReadScreen pay_freq, 1, 5, 64					'Finding the pay frequency'
					Select Case pay_freq							'Total monthly hours are determined by pay frequency specific multipliers'
					Case "1"										'Hours are added together as all earned income panels are checked'
						prosp_hrs = prosp_hrs
					Case "2"
						prosp_hrs = (2 * prosp_hrs)
					Case "3"
						prosp_hrs = (2.15 * prosp_hrs)
					Case "4"
						prosp_hrs = (4.3 * prosp_hrs)
					End Select
					prospective_hours = prospective_hours + prosp_hrs
				ELSE
					jobs_end_dt = replace(jobs_end_dt, " ", "/")	'Also considers jobs with a future end date'
					IF DateDiff("D", date, jobs_end_dt) > 0 THEN
						'Going into the PIC for a job with an end date in the future
						CALL write_value_and_transmit("X", 19, 38)
						EMReadScreen prosp_monthly, 8, 18, 56
						prosp_monthly = trim(prosp_monthly)
						IF prosp_monthly = "" THEN prosp_monthly = 0
						prosp_inc = prosp_inc + prosp_monthly
						EMReadScreen prosp_hrs, 8, 16, 50
						IF prosp_hrs = "        " THEN prosp_hrs = 0
						prosp_hrs = prosp_hrs * 1						'Added to ensure that prosp_hrs is a numeric
						EMReadScreen pay_freq, 1, 5, 64
						Select Case pay_freq
						Case "1"
							prosp_hrs = prosp_hrs
						Case "2"
							prosp_hrs = (2 * prosp_hrs)
						Case "3"
							prosp_hrs = (2.15 * prosp_hrs)
						Case "4"
							prosp_hrs = (4.3 * prosp_hrs)
						End Select
						'added seperate incremental variable to account for multiple jobs
						prospective_hours = prospective_hours + prosp_hrs
					END IF
				END IF
				transmit
				EMReadScreen JOBS_panel_current, 1, 2, 73
				'looping until all the jobs panels are calculated
				If cint(JOBS_panel_current) < cint(num_of_JOBS) then transmit
			Loop until cint(JOBS_panel_current) = cint(num_of_JOBS)
		End If
		'BUSI'
		EMWriteScreen "BUSI", 20, 71			'Checkin BUSI for earned income in self employment'
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)	'Looking for the reported client only'
		EMReadScreen num_of_BUSI, 1, 2, 78
		IF num_of_BUSI <> "0" THEN				'If any BUSI panels exist for this client - script will check budgeted hours'
			DO
				EMReadScreen busi_end_dt, 8, 5, 72		'Looking for BUSI income with no end'
				busi_end_dt = replace(busi_end_dt, " ", "/")
				IF IsDate(busi_end_dt) = True THEN
					IF DateDiff("D", date, busi_end_dt) > 0 THEN
						EMReadScreen busi_inc, 8, 10, 69
						busi_inc = trim(busi_inc)
						EMReadScreen busi_hrs, 3, 13, 74
						busi_hrs = trim(busi_hrs)
						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0	'Adding hours and income to any others found'
						prosp_inc = prosp_inc + busi_inc
						prosp_hrs = prosp_hrs + busi_hrs
						prospective_hours = prospective_hours + busi_hrs
					END IF
				ELSE
					IF busi_end_dt = "__/__/__" THEN
						EMReadScreen busi_inc, 8, 10, 69
						busi_inc = trim(busi_inc)
						EMReadScreen busi_hrs, 3, 13, 74
						busi_hrs = trim(busi_hrs)
						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
						prosp_inc = prosp_inc + busi_inc
						prosp_hrs = prosp_hrs + busi_hrs
						prospective_hours = prospective_hours + busi_hrs
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		END IF		'All of the budgteted earned income has been gathered and added'
		IF prosp_inc >= 935.25 OR prospective_hours >= 129 THEN		'Clients working the equivalent of 30 hours/wk by hours or income are FSET exempt'
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to be earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions. | "
		ELSEIF prospective_hours >= 80 AND prospective_hours < 129 THEN		'Clients working at least 80 hours in 1 month are ABAWD exempt'
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption. | "
		END IF
		'UNEA'
		call navigate_to_MAXIS_screen ("STAT", "UNEA")		'Checking for UI income'
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_UNEA, 1, 2, 78
		IF num_of_UNEA <> "0" THEN
			DO
				EMReadScreen unea_type, 2, 5, 37
				EMReadScreen unea_end_dt, 8, 7, 68
				unea_end_dt = replace(unea_end_dt, " ", "/")	'Looking for end dates'
				IF IsDate(unea_end_dt) = True THEN
					IF DateDiff("D", date, unea_end_dt) > 0 THEN
						IF unea_type = "14" THEN				'If there is a UNEA panel for UI clt is FSET Exempt'
							Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
							Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions. | "
						End If
					END IF
				ELSE
					IF unea_end_dt = "__/__/__" THEN
						IF unea_type = "14" THEN
						 	Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
							Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions. | "
						End If
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		End If
		'PBEN'
		EMWriteScreen "PBEN", 20, 71		'Going to PBEN for other exemptions'
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)	'Looking for only PBEN panels for the reported client'
		EMReadScreen num_of_PBEN, 1, 2, 78
		IF num_of_PBEN <> "0" THEN
			pben_row = 8
			DO
				EMReadScreen pben_type, 2, pben_row, 24
				IF pben_type = "02" THEN			'SSI pending'
					EMReadScreen pben_disp, 1, pben_row, 77
					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have pending, appealing, or eligible SSI benefits. Please review for ABAWD and SNAP E&T exemption. | "
						EXIT DO
					ELSE
						pben_row = pben_row + 1
					END IF
				ELSEIF pben_type = "12" THEN		'UI pending'
					EMReadScreen pben_disp, 1, pben_row, 77
					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption. | "
						EXIT DO
					ELSE
						pben_row = pben_row + 1
					END IF
				ELSE
					pben_row = pben_row + 1	'Needs to check all of the PBEN rows as SSI/UI may not be first on the list'
				END IF
			LOOP UNTIL pben_row = 14
		END IF
		'PREG'
		CALL navigate_to_MAXIS_screen("STAT", "PREG")	'Looks for a PREG panel'
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)	'For the reported client only'
		EMReadScreen num_of_PREG, 1, 2, 78
		EMReadScreen preg_end_dt, 8, 12, 53
		IF num_of_PREG <> "0" AND preg_end_dt <> "__ __ __" THEN		'With no end date'
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have active pregnancy. Please review for ABAWD exemption. | "
		END IF
		'SCHL/STIN/STEC
		CALL navigate_to_MAXIS_screen("STAT", "SCHL")		'Going to school'
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)	'For the reported client only'
		EMReadScreen num_of_SCHL, 1, 2, 78
		IF num_of_SCHL = "1" THEN
			EMReadScreen school_status, 1, 6, 40
			IF school_status <> "N" THEN
				Banked_Month_Client_Array(send_to_DHS, item) = FALSE	'Removing this client from DHS report - reason on next line'
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions. | "
			End If
		END IF

		'//////////WREG PORTION//////////////////////////////////////////////
		'This is intense, the script is going to check every line on the WREG tracker to list all of the counted ABAWD months for the report'

		Call navigate_to_MAXIS_screen("stat","wreg")		'navigates to stat/wreg
		EMWriteScreen Banked_Month_Client_Array(memb_num, item), 20, 76
		transmit
		EMReadScreen wreg_code,  2, 8,  50
		EMReadScreen abawd_code, 2, 13, 50
		IF wreg_code <> "30" Then	'ALL Banked Month clients should have WREG coded 30-13'
			Banked_Month_Client_Array(send_to_DHS,     item) = FALSE	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "WREG code is not a 30 (Mandatory FSET participant). Review. | "
		Elseif abawd_code <> "13" then 			'this is to make sure that 30/11 (second set cases) are removed from the report for the report month
			Banked_Month_Client_Array(send_to_DHS,     item) = FALSE	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "ABAWD code is not a 13 (ABAWD banked months). Review. | "
		END if

		report_date = MAXIS_footer_month & "/" & MAXIS_footer_year			'creating date variables to measure against person note counted dates

		EMReadScreen wreg_total, 1, 2, 78
		IF wreg_total <> "0" THEN
			EmWriteScreen "x", 13, 57		'Pulls up the WREG tracker'
			transmit
			EMREADScreen tracking_record_check, 15, 4, 40  		'adds cases to the rejection list if the ABAWD tracking record cannot be accessed.
			If tracking_record_check <> "Tracking Record" then
				Banked_Month_Client_Array(send_to_DHS,     item) = FALSE	'Removing this client from DHS report - reason on next line'
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Unable to access the ABAWD tracking record. Review manually. | "
			ELSE
				bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
				bene_yr_row = 10
				abawd_counted_months = 0					'delclares the variables values at 0
				second_abawd_period = 0
				month_count = 0
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
					abawd_counted_months_string = counted_date_month & "/" & counted_date_year

					'reading to see if a month is counted month or not
					EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col

					'rejects cases that do not have the report month coded as a counted month
					If report_date = abawd_counted_months_string then
						if is_counted_month = "X" then
							if is_counted_month = "M" then
								Banked_Month_Client_Array(send_to_DHS,     item) = FALSE	'Removing this client from DHS report - reason on next line'
								Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "ABAWD tracking record is coded as a counted ABAWD month (codes X or M) for " & report_date & ". Review manually. | "
							END IF
						END IF
					END IF

					'counting and checking for counted ABAWD months
					IF is_counted_month = "X" or is_counted_month = "M" THEN
						If abawd_counted_months_string <> report_date then
							EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
							abawd_counted_months_string = counted_date_month & "/" & counted_date_year
							abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
							abawd_counted_months = abawd_counted_months + 1				'adding counted months
						END IF
					END IF

					'declaring & splitting the abawd months array
					If left(abawd_info_list, 1) = "," then abawd_info_list = right(abawd_info_list, len(abawd_info_list) - 1)
					abawd_months_array = Split(abawd_info_list, ",")

					'counting and checking for second set of ABAWD months
					IF is_counted_month = "Y" or is_counted_month = "N" THEN
						EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
						second_abawd_period = second_abawd_period + 1				'adding counted months
						second_counted_months_string = counted_date_month & "/" & counted_date_year			'creating new variable for array
						second_set_info_list = second_set_info_list & ", " & second_counted_months_string	'adding variable to list to add to array
					END IF

					'declaring & splitting the second set of abawd months array
					If left(second_set_info_list, 1) = "," then second_set_info_list = right(second_set_info_list, len(second_set_info_list) - 1)
					second_months_array = Split(second_set_info_list,",")

					bene_mo_col = bene_mo_col - 4		're-establishing serach once the end of the row is reached
					IF bene_mo_col = 15 THEN
						bene_yr_row = bene_yr_row - 1
						bene_mo_col = 63
					END IF
					month_count = month_count + 1
				LOOP until month_count = 36
			PF3
			End if
		END If
		'END OF ABAWD MONTHS AND SECOND ABAWD MONTHS----------------------------------------------------------------------------------------------------

		'Reading the person notes regarding which months are counted as banked months
		PF5			'navigates to Person note from WREG PANEL
		'adds case to the rejected list if cannot access the person notes screen. This is usually for INACTIVE cases or out-of-county cases.
		EMReadScreen person_note_confirmation, 12, 2, 31
		If person_note_confirmation <> "Person Notes" then
			Banked_Month_Client_Array(send_to_DHS, item) = False	'Removing this client from DHS report - reason on next line'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Unable to determine counted banked months, no access to person notes. Review manually. | "
		ELSE
			DO
				PNOTE_row = 5		'establishes the row to start searching the Person notes from
				Do
					EMReadScreen counted_banked_month, 12, PNOTE_row, 31
					If counted_banked_month = "            " then exit do 'if blank then stops checking
					If counted_banked_month = "Banked Month" then
						EMReadScreen abawd_counted_months_string, 5, PNOTE_row, 49
						If abawd_counted_months_string < report_date then banked_months_list = banked_months_list & abawd_counted_months_string & ", "  'does not add dates that are report month or later to the array
					END IF
					PNOTE_row = PNOTE_row + 1
				LOOP until PNOTE_row = 18
				PF8
				EMReadScreen notes_exist, 1, 5, 3
				EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
			Loop until last_page_check = "THIS IS THE LAST PAGE" OR notes_exist <> "_"

			Dim PNOTE_array
			Dim Filter_array

			'declaring & splitting for the person note cases
			banked_months_list = trim(banked_months_list)
			if right(banked_months_list, 1) = "," then banked_months_list = left(banked_months_list, len(banked_months_list) - 1)
			'created new array of the banked months list cases
			PNOTE_array = Split(banked_months_list, ",")

			For each PNOTE in PNOTE_array	'This will remove any counted month that was actually a banked month'
				Filter_array = Filter(abawd_months_array, PNOTE, False, 1) 'The value of 1 is vbTextCompare - which will perform a textual comparison between the PNOTE month and the elements in the abawd_months_array
				abawd_counted_months = abawd_counted_months - 1				'subtracts counted months
				abawd_months_array = Filter_array						'establishing the values of both arrays are the same so that the PNOTE month that was removed stays removed from array
			NEXT

			'Now all the information about the counted months will be added to the array'
			Banked_Month_Client_Array(abawd_count,       item) = abawd_counted_months
			Banked_Month_Client_Array(second_count,      item) = second_abawd_period
			Banked_Month_Client_Array(abawd_used,        item) = Join(abawd_months_array, ", ")
			Banked_Month_Client_Array(second_abawd_used, item) = Join(second_months_array, ", ")
			If Banked_Month_Client_Array(second_abawd_used, item) = "" Then Banked_Month_Client_Array(second_abawd_used, item) = "None"	'If this array is blank - added none so there is no blank on the DHS report'

			IF Banked_Month_Client_Array(abawd_count, item) < 3 Then
				Banked_Month_Client_Array(send_to_DHS, item) = False	'Removing this client from DHS report - reason on next line'
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client has a WREG panel coded with fewer than 3 counted regular ABAWD months. | "
			End If

			'Write a new person note only for cases that are being sent to DHS on the 'true' list
			If (developer_mode_checkbox = unchecked AND Banked_Month_Client_Array(send_to_DHS, item) = True) then
				PF5		'enters person note screen
				'adds case to the rejected list if cannot person note
				EMReadScreen person_note_confirmation, 12, 2, 31
				If person_note_confirmation <> "Person Notes" then
					Banked_Month_Client_Array(send_to_DHS, item) = False	'Removing this client from DHS report - reason on next line'
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Unable to person note this case. Case may be in another county. | "
				ELSE
					'if not person not exists, person note goes directly into edit mode
					EMreadscreen edit_mode_required_check, 6, 5, 3
					If edit_mode_required_check = "      " then
						EMWriteScreen "Banked Month Used " & report_date, 5, 3
						EMWriteScreen "Case has been counted and reported to DHS.", 6, 3
					ElseIF edit_mode_required_check <> "      " then
						'creating a Do loop to ensure that duplicate person notes are not being made
						PNOTE_row = 5		'establishes the row to start searching the Person notes from
						Do
							EMReadScreen counted_banked_month, 12, PNOTE_row, 31
							If counted_banked_month = "Banked Month" then EMReadScreen abawd_counted_months_string, 5, PNOTE_row, 49
							If abawd_counted_months_string = report_date then exit do	'if person note has already been made for the report date, then does not person note
							PNOTE_row = PNOTE_row + 1	'adds incremental to row to search
						LOOP until PNOTE_row = 18
						If PNOTE_row = 18 then
							PF9
							EMWriteScreen "Banked Month Used " & report_date, 5, 3
							EMWriteScreen "Case has been counted and reported to DHS.", 6, 3
						END IF
					END IF
				END IF
			END If
			PF3 'exits person note'

			'clears values of the following variables
			abawd_counted_months_string = ""
			abawd_info_list = ""
			second_counted_months_string = ""
			second_set_info_list = ""
			banked_months_list = ""
			abawd_counted_months = ""
			second_abawd_period = ""
		End If
	END If
Next

'-----------------------------------END OF WREG PIECE---------------------------------------------------------------
'Dialog to select the file that users will send to DHS
BeginDialog DHS_Report_Dialog, 0, 0, 226, 65, "DHS Banked Months"
  EditBox 10, 20, 160, 15, DHS_Banked_Month_Report_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 175, 20, 45, 15, "Browse...", select_a_file_button
    OkButton 115, 45, 50, 15
    CancelButton 170, 45, 50, 15
  Text 5, 5, 215, 10, "Select the Excel File that you use to report Banked Months to DHS"
EndDialog

'Runs the dialog
Do
	Dialog DHS_Report_Dialog
	cancel_confirmation
	Call File_Selection_System_Dialog(DHS_Banked_Month_Report_excel_file_path, ".xlsx")  'References the function above to have the user seach for their file'
	call excel_open(DHS_Banked_Month_Report_excel_file_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
Loop until DHS_Banked_Month_Report_excel_file_path <> ""

'This reads every worksheet name in the selected excel file and creates an array that will be used to determine which month is being reported'
For Each objWorkSheet In objWorkbook.Worksheets
	DHS_report_month_list = DHS_report_month_list & "~" & objWorkSheet.Name
	If left(DHS_report_month_list, 1) = "~" then DHS_report_month_list = right(DHS_report_month_list, len(DHS_report_month_list) - 1)
	DHS_report_month_array = Split(DHS_report_month_list,"~")
Next

'The user already selected the month in the initial excel sheet - this was used to set the footer month.'
'Now the footer month is used to select the right worksheet in the DHS report to match'
Select Case MAXIS_footer_month
Case "01"
	DHS_report_month = DHS_report_month_array(0)	'January, arrays start at 0'
Case "02"
	DHS_report_month = DHS_report_month_array(1)	'February'
Case "03"
	DHS_report_month = DHS_report_month_array(2)	'March'
Case "04"
	DHS_report_month = DHS_report_month_array(3)	'April'
Case "05"
	DHS_report_month = DHS_report_month_array(4)	'May'
Case "06"
	DHS_report_month = DHS_report_month_array(5)	'June'
Case "07"
	DHS_report_month = DHS_report_month_array(6)	'July'
Case "08"
	DHS_report_month = DHS_report_month_array(7)	'August'
Case "09"
	DHS_report_month = DHS_report_month_array(8)	'September'
Case "10"
	DHS_report_month = DHS_report_month_array(9)	'October'
Case "11"
	DHS_report_month = DHS_report_month_array(10)	'November'
Case "12"
	DHS_report_month = DHS_report_month_array(11)	'December'
End Select

'Activates the selected worksheet'
objExcel.worksheets(DHS_report_month).Activate

excel_row = 2
abawd_count_range = 1
second_count_range = 1

'Excel Column Constants'
Const               county_column = 1'
Const          case_number_column = 2'
Const           PMI_number_column = 3'
Const counted_ABAWD_months_column = 4'
Const  second_three_months_column = 5'
Const         WREG_updated_column = 6'
Const   total_ABAWD_Months_column = 7'
Const     total_second_Set_column = 8'
Const             comments_column = 9'

'All of the information has been stored in an array and now needs to be entered into a spreadsheet. Each line of the spreadsheet will be one array entry'
For clients_to_report = 0 to UBound(Banked_Month_Client_Array,2)
	IF Banked_Month_Client_Array(send_to_DHS, clients_to_report) = TRUE Then
		objExcel.Cells(excel_row,              county_column).Value = county_name
		objExcel.Cells(excel_row,         case_number_column).Value = Banked_Month_Client_Array (case_num,          clients_to_report)	'Adding the case number'
		objExcel.Cells(excel_row,          PMI_number_column).Value = Banked_Month_Client_Array (clt_pmi,           clients_to_report)	'Adding the PMI number'
		objExcel.Cells(excel_row,counted_ABAWD_months_column).Value = Banked_Month_Client_Array (abawd_used,        clients_to_report)	'Adding the list of ABAWD months used'
		objExcel.Cells(excel_row, second_three_months_column).Value = Banked_Month_Client_Array (second_abawd_used, clients_to_report)	'Adding the list of Second ABAWD months used'
		objExcel.Cells(excel_row,        WREG_updated_column).Value = "Yes"		'Hard coded because if this was not coded correctly the case would not be added'
		objExcel.Cells(excel_row,  total_ABAWD_Months_column).Value = Banked_Month_Client_Array (abawd_count,       clients_to_report)	'Adding the total of ABAWD months used'
		objExcel.Cells(excel_row,    total_second_Set_column).Value = Banked_Month_Client_Array (second_count,      clients_to_report)
		objExcel.Cells(excel_row,            comments_column).Value = Banked_Month_Client_Array (comments,          clients_to_report)	'Adding any comments that were on the initial spreadsheet'
		excel_row = excel_row + 1	'Goes to the next Excel row for the next itteration of the For..Next'
		abawd_total_range = abawd_count_range + 1		'Setting the area to do math in Excel'
		second_count_range = second_count_range + 1
	ElseIf Banked_Month_Client_Array(send_to_DHS,clients_to_report) = FALSE Then
		need_word_doc = TRUE	'If every client on the list is added to the DHS report - no rejected list is needed - this variable sets if the script will create a rejected list'
	End If
Next

excel_row = 2
If need_word_doc  = TRUE Then		'This will create the second report of cases rejected'
	Set objNewExcel = CreateObject("Excel.Application")
	Set objWorkbook = objNewExcel.Workbooks.Add()
	objNewExcel.Cells(1, 1).Value = "CASE NUMBER"
	objNewExcel.Cells(1, 1).Font.Bold = True
	objNewExcel.Cells(1, 2).Value = "FIRST NAME"
	objNewExcel.Cells(1, 2).Font.Bold = True
	objNewExcel.Cells(1, 3).Value = "LAST NAME"
	objNewExcel.Cells(1, 3).Font.Bold = True
	objNewExcel.Cells(1, 4).Value = "Reason not reported to DHS"
	objNewExcel.Cells(1, 4).Font.Bold = True
	For	not_reported_clients = 0 to UBound(Banked_Month_Client_Array,2)
		IF Banked_Month_Client_Array(send_to_DHS,not_reported_clients) = False Then		'Only the entries that were not on the DHS report will be on this report'
			objNewExcel.Cells(excel_row, 1).Value = Banked_Month_Client_Array (case_num,        not_reported_clients)	'Adding case number'
			objNewExcel.Cells(excel_row, 2).Value = Banked_Month_Client_Array (clt_first_name,  not_reported_clients)	'Adding client name'
			objNewExcel.Cells(excel_row, 3).Value = Banked_Month_Client_Array (clt_last_name,   not_reported_clients)
			objNewExcel.Cells(excel_row, 4).Value = Banked_Month_Client_Array (reason_excluded, not_reported_clients)	'Adding the list of reasons the client is not on the DHS report'
			excel_row = excel_row + 1
		End If
	Next
End If
'Formatting the spreadsheet so it looks good'
objNewExcel.columns(4).WrapText = True
For col_to_autofit = 1 to 4
	ObjNewExcel.columns(col_to_autofit).AutoFit()
Next
'objNewExcel.columns(4).columnwidth = 850
objNewExcel.Visible = True

STATS_counter = STATS_counter - 1 					'removing 1 count from stats counter as we start with 1, for accurate count
script_end_procedure("Success!")
