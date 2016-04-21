'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - MONTHLY BANKED MONTHS DATA GATHER.vbs"
start_time = timer

STATS_counter = 1                          'sets the stats counter at one
'STATS_manualtime = ***                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'FUNCTIONS that are currently not in the FuncLib that are used in this script----------------------------------------------------------------------------------------------------
'Veronicas function that allows the user to search for a local file instead of having the file location hard coded into the script'
'This can be removed as the function is in FuncLib'
Function File_Selection_System_Dialog(file_selected)
    'Creates a Windows Script Host object
    Set wShell=CreateObject("WScript.Shell")

    'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
    Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

    'Creates the file_selected variable from the exit
    file_selected = oExec.StdOut.ReadLine
End function

EMConnect ""		'connecting to MAXIS

'Dialog needed here before pushing to master'
MsgBox "You need to open the Excel File that has the list of clients reported as using a banked month for the month being reported." & _
  VBNewLine & VBNewLine & "Be sure your spreadsheet is in the correct format." 'Notice to the user that a finder window will open for them to search for their list of client that have used banked months'
Call File_Selection_System_Dialog(list_reported_banked_month_clients)  'References the function above to have the user seach for their file'
call excel_open(list_reported_banked_month_clients, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

'This reads every worksheet name in the selected excel file and creates a list for the drop down in the dialog'
For Each objWorkSheet In objWorkbook.Worksheets
	month_list = month_list & chr(9) & objWorkSheet.Name
Next

'This is the dialog for the user to select which month (or worksheet) of data they are going to generate a report for'
BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 211, 70, "SNAP Banked Month Reporting Dialog"
  DropListBox 65, 25, 140, 15, "select one..." & month_list, report_month_dropdown
  ButtonGroup ButtonPressed
	OkButton 100, 45, 50, 15
	CancelButton 155, 45, 50, 15
  Text 5, 10, 190, 10, "Select the month that you are creating the report for."
  Text 5, 30, 55, 10, "Month to Report:"
EndDialog

'Runs the dialog'
Do
	Dialog SNAP_Banked_Month_Report_Dialog
	cancel_confirmation
Loop until report_month_dropdown <> "select one..."

'Activates the selected worksheet'
objExcel.worksheets(report_month_dropdown).Activate

'This assigns a footer month and year based on the worksheet names selected in the dropdown from the dialog'
Select Case report_month_dropdown
Case "January 2016"
	footer_month = "01"
	footer_year = "16"
Case "February 2016"
	footer_month = "02"
	footer_year = "16"
Case "March 2016"
	footer_month = "03"
	footer_year = "16"
Case "April 2016"
	footer_month = "04"
	footer_year = "16"
Case "May 2016"
	footer_month = "05"
	footer_year = "16"
Case "June 2016"
	footer_month = "06"
	footer_year = "16"
Case "July 2016"
	footer_month = "07"
	footer_year = "16"
Case "August 2016"
	footer_month = "08"
	footer_year = "16"
Case "September 2016"
	footer_month = "09"
	footer_year = "16"
Case "October 2016"
	footer_month = "10"
	footer_year = "16"
Case "November 2016"
	footer_month = "11"
	footer_year = "16"
Case "December 2016"
	footer_month = "12"
	footer_year = "16"
End Select

'Sets up the array to store all the information for each client'
Dim Banked_Month_Client_Array ()
ReDim Banked_Month_Client_Array (14, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num          = 1
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
Const send_to_DHS       = 12
Const reason_excluded   = 13
Const clt_filter        = 14

'Now the script adds all the clients on the excel list into an array
excel_row = 3 're-establishing the row to start checking the members for
entry_record = 0
Do                                                                                          'Loops until there are no more cases in the Excel list
	case_number = objExcel.cells(excel_row, 4).Value          're-establishing the case numbers
	If case_number = "" then exit do
	case_number = trim(case_number)
	client_first_name = objExcel.cells(excel_row, 3).Value
	client_last_name  = objExcel.cells(excel_row, 2).Value             're-establishing the client name
	client_first_name = UCase(trim(client_first_name))
	client_last_name  = UCase(trim(client_last_name))
'Adding client information to the array'
	ReDim Preserve Banked_Month_Client_Array(14, entry_record)
	Banked_Month_Client_Array (case_num,       entry_record) = case_number
	Banked_Month_Client_Array (clt_last_name,  entry_record) = client_last_name
	Banked_Month_Client_Array (clt_first_name, entry_record) = client_first_name
	Banked_Month_Client_Array (clt_name,       entry_record) = client_first_name & " " & client_last_name
	Banked_Month_Client_Array (comments,       entry_record) = objExcel.cells(excel_row, 6).Value
	Banked_Month_Client_Array (send_to_DHS,    entry_record) = TRUE

	'MsgBox client_first_name & " " & client_last_name & VBNewLine & Banked_Month_Client_Array (clt_name,entry_record)
	entry_record = entry_record + 1
	excel_row = excel_row + 1
Loop

'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document
'and don't want the script to be confused
objExcel.Quit

'Now we will get PMI and Member Number for each client on the array.'
For item = 0 to UBound(Banked_Month_Client_Array, 2)
	case_number = Banked_Month_Client_Array(case_num,item)	'Case number is set for each loop as it is used in the FuncLib functions'
	Call navigate_to_MAXIS_screen("INFC", "WORK")						'Finding client information on STAT MEMB'
	EMReadScreen WORK_check, 4, 2, 51
	IF WORK_check = "WORK" Then
		work_maxis_row = 7
		DO
			EMReadScreen client_referred, 26, work_maxis_row, 7
			memb_check = MsgBox ("Client listed on your report: " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name, item) & _
			  vbNewLine &        "Client name listed in MAXIS: " & trim(client_referred) & vbNewLine & vbNewLine & "Is this the client you are reporting as using banked months?", vbYesNo + vbQuestion, "Confirm Client using Banked Monhts")
			If memb_check = vbYes Then
				EMReadScreen Banked_Month_Client_Array(clt_pmi,  item), 8, work_maxis_row, 34
				EMReadScreen Banked_Month_Client_Array(memb_num, item), 2, work_maxis_row, 3
			ElseIf memb_check = vbNo Then
				EMReadScreen next_clt, 1, (work_maxis_row + 1), 7
				If next_clt = " " Then
					MsgBox "There are no additional clients on this case that have had a workforce referral. Since banked months require E&T participation, there must be a referral This client - " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name, item) & _
				    " - will not be added to the DHS report."
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "Person not matched with name in MAXIS. | "
				End If
			End If
			work_maxis_row = work_maxis_row + 1
		Loop until next_clt = " " OR memb_check = vbYes
	Else
		Banked_Month_Client_Array(send_to_DHS, item) = FALSE
		Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "No Workforce1 referral was done. Banked Months requires client to participate in E&T, so a Workforce 1 Referral needs to be completed. | "
	End If
	If Banked_Month_Client_Array(send_to_DHS, item) = TRUE Then
		call navigate_to_MAXIS_screen ("ELIG", "FS")
		EMReadScreen fs_version, 8, 3, 3
		If fs_version = "UNAPPROV" Then
			EMReadScreen vers_number, 1, 2, 19
			If vers_number = "1" Then
				'MsgBox "No approved version of SNAP exists for this case in the given month. This client - " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name,item) & " - will not be added to the DHS Report"
				Banked_Month_Client_Array(send_to_DHS, item) = FALSE
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "SNAP not approved in " & footer_month & "/" & footer_year & " | "
			End If
			EMWriteScreen "0" & (abs(vers_number) - 1), 19, 78
			transmit
		ElseIf fs_version = "        " Then
			'MsgBox "No version of SNAP exists for this case in the given month. This client - " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name,item) & " - will not be added to the DHS Report"
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "No ELIG/FS version exists in " & footer_month & "/" & footer_year & " | "
		End If
		elig_maxis_row = 7
		Do
			EMReadScreen clt_on_snap, 2, elig_maxis_row, 10
			IF clt_on_snap = Banked_Month_Client_Array(memb_num,item) Then
				EMReadScreen pers_elig, 8, elig_maxis_row, 57
				IF pers_elig <> "ELIGIBLE" Then
					'MsgBox "This client - " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name,item) & " - is not listing as eligible for SNAP and will not be added to the DHS Report"
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "Client listed as Ineligible for SNAP on ELIG/FS for " & footer_month & "/" & footer_year & " | "
				End If
			ElseIf clt_on_snap = "  " Then
				'MsgBox "This client - " & Banked_Month_Client_Array(clt_last_name, item) & ", " & Banked_Month_Client_Array(clt_first_name,item) & " - could not be found on the SNAP Eligibility and will not be added to the DHS Report"
				Banked_Month_Client_Array(send_to_DHS, item) = FALSE
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "Client not listed on ELIG/FS for " &  footer_month & "/" & footer_year & " | "
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
		EMReadScreen fs_prorated, 8, 11,40
		IF fs_prorated = "Prorated" Then
			'MsgBox "SNAP is prorated in this month for case # " & Banked_Month_Client_Array(case_num,item) & ". This case will not be reported to DHS as using a Banked Month."
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded,item) & "SNAP is prorated in " & footer_month & "/" & footer_year & " | "
		End If
		'///////SCRIPT WILL NOW CHECK FOR POSSIBLE EXPEMTIONS FOR CLIENT'
		'Age exemption'
		call navigate_to_MAXIS_screen ("STAT", "MEMB")																							'Cient age is listed on STAT MEMB'
		Call write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)			'Writes the clt reference number on the command line to get to the correct MEMB Panel'
		EMReadScreen cl_age, 2, 8, 76																													'Reads the client age'
		cl_age = abs(cl_age)																																	'Makes sure the age is seen in the script as a number for the math that come next'
		IF cl_age < 18 OR cl_age >= 50 THEN 																									'Compares to the age exclusions'
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE 																'Codes the array to not include this client and the reason'
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have exemption. Age = " & cl_age & ". | "
		End If
		'Exemptions for Disability'
		disa_status = false																																		'Resets the variable for the looping'
		call navigate_to_MAXIS_screen("STAT", "DISA")																							'Information about disability is on STAT/DISA'
		Call write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)				'Enters the clt reference number to get to the right disa panel'
		EMReadScreen num_of_DISA, 1, 2, 78
		IF num_of_DISA <> "0" THEN
			EMReadScreen disa_end_dt, 10, 6, 69
			disa_end_dt = replace(disa_end_dt, " ", "/")
			EMReadScreen cert_end_dt, 10, 7, 69
			cert_end_dt = replace(cert_end_dt, " ", "/")
			IF IsDate(disa_end_dt) = True THEN
				IF DateDiff("D", date, disa_end_dt) > 0 THEN
					disa_status = True
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA end date = " & disa_end_dt & ". | "
				END IF
			ELSE
				IF disa_end_dt = "__/__/____" OR disa_end_dt = "99/99/9999" THEN
					disa_status = True
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA has no end date. | "
				END IF
			END IF
			IF IsDate(cert_end_dt) = True AND disa_status = False THEN
				IF DateDiff("D", date, cert_end_dt) > 0 THEN
					Banked_Month_Client_Array(send_to_DHS, item) = FALSE
					Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA Certification end date = " & cert_end_dt & ". | "
				End If
			ELSE
				IF cert_end_dt = "__/__/____" OR cert_end_dt = "99/99/9999" THEN
					EMReadScreen cert_begin_dt, 8, 7, 47
					IF cert_begin_dt <> "__ __ __" THEN
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have disability exemption. DISA certification has no end date. | "
					End If
				END IF
			END IF
		End If
		'Checking for earned income exemptions'
		'JOBS'
		prosp_inc = 0
		prosp_hrs = 0

		CALL navigate_to_MAXIS_screen("STAT", "JOBS")
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_JOBS, 1, 2, 78
		IF num_of_JOBS <> "0" THEN
			DO
				EMReadScreen jobs_end_dt, 8, 9, 49
				EMReadScreen cont_end_dt, 8, 9, 73
				IF jobs_end_dt = "__ __ __" THEN
					CALL write_value_and_transmit("X", 19, 38)
					EMReadScreen prosp_monthly, 8, 18, 56
					prosp_monthly = trim(prosp_monthly)
					IF prosp_monthly = "" THEN prosp_monthly = 0
					prosp_inc = prosp_inc + prosp_monthly
					EMReadScreen pp_hrs, 8, 16, 50
					IF pp_hrs = "        " THEN pp_hrs = 0
					pp_hrs = abs(pp_hrs)
					EMReadScreen pay_freq, 1, 5, 64
					Select Case pay_freq
					Case "1"
						prosp_hrs = prosp_hrs + pp_hrs
					Case "2"
						prosp_hrs = prosp_hrs + pp_hrs * 2
					Case "3"
						prosp_hrs = prosp_hrs + pp_hrs * 2.15
					Case "4"
						prosp_hrs = prosp_hrs + pp_hrs * 4.3
					End Select
				ELSE
					jobs_end_dt = replace(jobs_end_dt, " ", "/")
					IF DateDiff("D", date, jobs_end_dt) > 0 THEN
						'Going into the PIC for a job with an end date in the future
						CALL write_value_and_transmit("X", 19, 38)
						EMReadScreen prosp_monthly, 8, 18, 56
						prosp_monthly = trim(prosp_monthly)
						IF prosp_monthly = "" THEN prosp_monthly = 0
						prosp_inc = prosp_inc + prosp_monthly
						EMReadScreen pp_hrs, 8, 16, 50
						IF pp_hrs = "        " THEN pp_hrs = 0
						EMReadScreen pay_freq, 1, 5, 64
						Select Case pay_freq
						Case "1"
							prosp_hrs = prosp_hrs + pp_hrs
						Case "2"
							prosp_hrs = prosp_hrs + pp_hrs * 2
						Case "3"
							prosp_hrs = prosp_hrs + pp_hrs * 2.15
						Case "4"
							prosp_hrs = prosp_hrs + pp_hrs * 4.3
						End Select
					END IF
				END IF
				transmit
				transmit
				EMReadScreen enter_a_valid_command, 13, 24, 2
			LOOP UNTIL enter_a_valid_command = "ENTER A VALID"
		End If
		'BUSI'
		EMWriteScreen "BUSI", 20, 71
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_BUSI, 1, 2, 78
		IF num_of_BUSI <> "0" THEN
			DO
				EMReadScreen busi_end_dt, 8, 5, 72
				busi_end_dt = replace(busi_end_dt, " ", "/")
				IF IsDate(busi_end_dt) = True THEN
					IF DateDiff("D", date, busi_end_dt) > 0 THEN
						EMReadScreen busi_inc, 8, 10, 69
						busi_inc = trim(busi_inc)
						EMReadScreen busi_hrs, 3, 13, 74
						busi_hrs = trim(busi_hrs)
						IF InStr(busi_hrs, "?") <> 0 THEN busi_hrs = 0
						prosp_inc = prosp_inc + busi_inc
						prosp_hrs = prosp_hrs + busi_hrs
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
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		END IF
		IF prosp_inc >= 935.25 OR prosp_hrs >= 129 THEN
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to be earning equivalent of 30 hours/wk at federal minimum wage. Please review for ABAWD and SNAP E&T exemptions. | "
		ELSEIF prosp_hrs >= 80 AND prosp_hrs < 129 THEN
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to be working at least 80 hours in the benefit month. Please review for ABAWD exemption. | "
		END IF
		'UNEA'
		call navigate_to_MAXIS_screen ("STAT", "UNEA")
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_UNEA, 1, 2, 78
		IF num_of_UNEA <> "0" THEN
			DO
				EMReadScreen unea_type, 2, 5, 37
				EMReadScreen unea_end_dt, 8, 7, 68
				unea_end_dt = replace(unea_end_dt, " ", "/")
				IF IsDate(unea_end_dt) = True THEN
					IF DateDiff("D", date, unea_end_dt) > 0 THEN
						IF unea_type = "14" THEN
							Banked_Month_Client_Array(send_to_DHS, item) = FALSE
							Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions. | "
						End If
					END IF
				ELSE
					IF unea_end_dt = "__/__/__" THEN
						IF unea_type = "14" THEN
						 	Banked_Month_Client_Array(send_to_DHS, item) = FALSE
							Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have active unemployment benefits. Please review for ABAWD and SNAP E&T exemptions. | "
						End If
					END IF
				END IF
				transmit
				EMReadScreen enter_a_valid, 13, 24, 2
			LOOP UNTIL enter_a_valid = "ENTER A VALID"
		End If
		'PBEN'
		EMWriteScreen "PBEN", 20, 71
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_PBEN, 1, 2, 78
		IF num_of_PBEN <> "0" THEN
			pben_row = 8
			DO
				EMReadScreen pben_type, 2, pben_row, 24
				IF pben_type = "02" THEN
					EMReadScreen pben_disp, 1, pben_row, 77
					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have pending, appealing, or eligible SSI benefits. Please review for ABAWD and SNAP E&T exemption. | "
						EXIT DO
					ELSE
						pben_row = pben_row + 1
					END IF
				ELSEIF pben_type = "12" THEN
					EMReadScreen pben_disp, 1, pben_row, 77
					IF pben_disp = "A" OR pben_disp = "E" OR pben_disp = "P" THEN
						Banked_Month_Client_Array(send_to_DHS, item) = FALSE
						Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have pending, appealing, or eligible Unemployment benefits. Please review for ABAWD and SNAP E&T exemption. | "
						EXIT DO
					ELSE
						pben_row = pben_row + 1
					END IF
				ELSE
					pben_row = pben_row + 1
				END IF
			LOOP UNTIL pben_row = 14
		END IF
		'PREG'
		CALL navigate_to_MAXIS_screen("STAT", "PREG")
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_PREG, 1, 2, 78
		EMReadScreen preg_end_dt, 8, 12, 53
		IF num_of_PREG <> "0" AND preg_end_dt <> "__ __ __" THEN
			Banked_Month_Client_Array(send_to_DHS, item) = FALSE
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to have active pregnancy. Please review for ABAWD exemption. | "
		END IF
		'SCHL/STIN/STEC
		CALL navigate_to_MAXIS_screen("STAT", "SCHL")
		CALL write_value_and_transmit(Banked_Month_Client_Array(memb_num, item), 20, 76)
		EMReadScreen num_of_SCHL, 1, 2, 78
		IF num_of_SCHL = "1" THEN
			EMReadScreen school_status, 1, 6, 40
			IF school_status <> "N" THEN
				Banked_Month_Client_Array(send_to_DHS, item) = FALSE
				Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client appears to be enrolled in school. Please review for ABAWD and SNAP E&T exemptions. | "
			End If
		END IF

		'------------NEWLY ADDED WREG PIECE-------------------------------------------------------------------------------
		'creating date variables to measure against person note counted dates
		report_date = footer_month & "/" & footer_year

		Call navigate_to_MAXIS_screen("stat","wreg")		'navigates to stat/wreg
		EMWriteScreen Banked_Month_Client_Array(memb_num, item), 20, 76
		transmit
		EMReadScreen wreg_code,  2, 8,  50
		EMReadScreen abawd_code, 2, 13, 50
		IF wreg_code <> "30" AND abawd_code <> "10" Then
			Banked_Month_Client_Array(send_to_DHS,     item) = FALSE
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "WREG is not coded 30-10. Review. | "
		End If
		EMReadScreen wreg_total, 1, 2, 78
		IF wreg_total <> "0" THEN
			EmWriteScreen "x", 13, 57
			transmit
			bene_mo_col = (15 + (4*cint(footer_month)))		'col to search starts at 15, increased by 4 for each footer month
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
					'reading to see if a month is counted month or not
					EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
					'counting and checking for counted ABAWD months
					IF is_counted_month = "X" or is_counted_month = "M" THEN
						EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
						abawd_counted_months_string = counted_date_month & "/" & counted_date_year
						If abawd_counted_months_string <> report_date then
							If counted_date_year < footer_year then 			'does not add dates that are report month or later to the array
								abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
								abawd_counted_months = abawd_counted_months + 1				'adding counted months
							Elseif abawd_counted_months_string < report_date then
								abawd_info_list = abawd_info_list & ", " & abawd_counted_months_string			'adding variable to list to add to array
								abawd_counted_months = abawd_counted_months + 1				'adding counted months
							END IF
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
						second_set_info_list = second_set_info_list & "," & second_counted_months_string	'adding variable to list to add to array
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
		END If
		'END OF ABAWD MONTHS AND SECOND ABAWD MONTHS----------------------------------------------------------------------------------------------------

		'Reading the person notes regarding which months are counted as banked months
		PF5			'navigates to Person note from WREG PANEL

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

		'declaring & splitting for the person note cases
		banked_months_list = trim(banked_months_list)
		if right(banked_months_list, 1) = "," then banked_months_list = left(banked_months_list, len(banked_months_list) - 1)
		'created new array of the banked months list cases
		PNOTE_array = Split(banked_months_list, ",")

		'msgbox "PNOTE array: " & Join(PNOTE_array, ",") & _
		'vbnewLine & "ABAWD array: " & Join(abawd_months_array, ",")

		Dim PNOTE_array
		Dim Filter_array

		For each PNOTE in PNOTE_array
			'msgbox "PNOTE: " & PNOTE
			Filter_array = Filter(abawd_months_array, PNOTE, False, 1) 'The value of 1 is vbTextCompare - which will perform a textual comparison between the PNOTE month and the elements in the abawd_months_array
			abawd_counted_months = abawd_counted_months - 1				'subtracts counted months
			abawd_months_array = Filter_array						'establishing the values of both arrays are the same so that the PNOTE month that was removed stays removed from array
			'msgbox "Filter array: " & Join(Filter_array, ",") & _
			'vbNewline & "abawd counted months" & abawd_counted_months
		NEXT

		Banked_Month_Client_Array(abawd_count,       item) = abawd_counted_months 'UBound(abawd_months_array) + 1
		Banked_Month_Client_Array(second_count,      item) = second_abawd_period  'UBound(second_months_array) + 1
		'If Filter_array <> "" Then Banked_Month_Client_Array(clt_filter,        item) = Join(Filter_array, ",")
		Banked_Month_Client_Array(abawd_used,        item) = Join(abawd_months_array, ", ")
		Banked_Month_Client_Array(second_abawd_used, item) = Join(second_months_array, ", ")
		If Banked_Month_Client_Array(second_abawd_used, item) = "" Then Banked_Month_Client_Array(second_abawd_used, item) = "None"

		IF Banked_Month_Client_Array(abawd_count, item) < 3 Then
			Banked_Month_Client_Array(send_to_DHS, item) = False
			Banked_Month_Client_Array(reason_excluded, item) = Banked_Month_Client_Array(reason_excluded, item) & "Client has a WREG panel coded with fewer than 3 counted regular ABAWD months. | "
		End If

		'Write a new person note'
		'Update WREG'

		PF3 	'exits Person Note

		abawd_counted_months_string = ""
		abawd_info_list = ""
		second_counted_months_string = ""
		second_set_info_list = ""

	End If
'-----------------------------------END OF WREG PIECE---------------------------------------------------------------
Next
'TESTING'
'For i = 0 to Ubound(Banked_Month_Client_Array,2)
''	MsgBox "Case # " & Banked_Month_Client_Array (case_num, i) & vbNewLine & "PMI: " & Banked_Month_Client_Array(clt_pmi,i) & _
''	  vbNewLine & "Memb " & Banked_Month_Client_Array(memb_num, i) & vbNewLine & "Name: " & Banked_Month_Client_Array(clt_name, i) & _
''	  vbNewLine & "Name again: " & Banked_Month_Client_Array (clt_first_name, i) & " " & Banked_Month_Client_Array(clt_last_name, i) & _
''	  vbNewLine & "Comments: " & Banked_Month_Client_Array(comments, i) & vbNewLine & "Counted ABAWD: " & Banked_Month_Client_Array(abawd_used, i) & _
''	  vbNewLine & "Second 3 Months: " & Banked_Month_Client_Array(second_abawd_used, i) & vbNewLine & "Filter: " & Banked_Month_Client_Array(clt_filter, i)
'Next


MsgBox "You need to open the Excel File that contains the DHS Banked Months Report" & _
  VBNewLine & VBNewLine & "Be sure your spreadsheet is in the correct format." 'Notice to the user that a finder window will open for them to search for their list of client that have used banked months'
Call File_Selection_System_Dialog(DHS_Banked_Month_Report)  'References the function above to have the user seach for their file'
call excel_open(DHS_Banked_Month_Report, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

'This reads every worksheet name in the selected excel file and creates a list for the drop down in the dialog'
For Each objWorkSheet In objWorkbook.Worksheets
	DHS_report_month_list = DHS_report_month_list & chr(9) & objWorkSheet.Name
Next

'Remove this option and have the script select it.'
'This is the dialog for the user to select which month (or worksheet) of data they are going to generate a report for'
BeginDialog DHS_Report_Dialog, 0, 0, 211, 70, "DHS Report Dialog"
  DropListBox 65, 25, 140, 15, "select one..." & DHS_report_month_list, DHS_report_dropdown
  ButtonGroup ButtonPressed
	OkButton 100, 45, 50, 15
	CancelButton 155, 45, 50, 15
  Text 5, 10, 190, 10, "Select which month you are reporting to DHS."
  Text 5, 30, 55, 10, "Month to Report:"
EndDialog

'Runs the dialog'
Do
	Dialog DHS_Report_Dialog
	cancel_confirmation
Loop until DHS_report_dropdown <> "select one..."

'Activates the selected worksheet'
objExcel.worksheets(DHS_report_dropdown).Activate

excel_row = 2

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

For clients_to_report = 0 to UBound(Banked_Month_Client_Array,2)
	IF Banked_Month_Client_Array(send_to_DHS, clients_to_report) = TRUE Then
		objExcel.Cells(excel_row,              county_column).Value = "Ramsey"
		objExcel.Cells(excel_row,         case_number_column).Value = Banked_Month_Client_Array (case_num,          clients_to_report)
		objExcel.Cells(excel_row,          PMI_number_column).Value = Banked_Month_Client_Array (clt_pmi,           clients_to_report)
		objExcel.Cells(excel_row,counted_ABAWD_months_column).Value = Banked_Month_Client_Array (abawd_used,        clients_to_report)
		objExcel.Cells(excel_row, second_three_months_column).Value = Banked_Month_Client_Array (second_abawd_used, clients_to_report)
		objExcel.Cells(excel_row,        WREG_updated_column).Value = "Yes"
		objExcel.Cells(excel_row,  total_ABAWD_Months_column).Value = Banked_Month_Client_Array (abawd_count,       clients_to_report)
		objExcel.Cells(excel_row,    total_second_Set_column).Value = Banked_Month_Client_Array (second_count,      clients_to_report)
		objExcel.Cells(excel_row,            comments_column).Value = Banked_Month_Client_Array (comments,          clients_to_report)
		excel_row = excel_row + 1
	ElseIf Banked_Month_Client_Array(send_to_DHS,clients_to_report) = FALSE Then
		need_word_doc = TRUE
	End If
Next

excel_row = 2
If need_word_doc  = TRUE Then
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
		IF Banked_Month_Client_Array(send_to_DHS,not_reported_clients) = False Then
			objNewExcel.Cells(excel_row, 1).Value = Banked_Month_Client_Array (case_num,        not_reported_clients)
			objNewExcel.Cells(excel_row, 2).Value = Banked_Month_Client_Array (clt_first_name,  not_reported_clients)
			objNewExcel.Cells(excel_row, 3).Value = Banked_Month_Client_Array (clt_last_name,   not_reported_clients)
			objNewExcel.Cells(excel_row, 4).Value = Banked_Month_Client_Array (reason_excluded, not_reported_clients)
			excel_row = excel_row + 1
		End If
	Next
End If
objNewExcel.columns(4).WrapText = True
For col_to_autofit = 1 to 4
	ObjNewExcel.columns(col_to_autofit).AutoFit()
Next
'objNewExcel.columns(4).columnwidth = 850
objNewExcel.Visible = True

'If need_word_doc  = TRUE Then
''	Set objBlank
''	Set objWord = CreateObject("Word.Application")
''	Set objDoc = objWord.Documents.Add()
''	Set objSelection = objWord.Selection
''	objSelection.Font.Name = "Ariel"
''	objSelection.Font.Size = "12"
''	For	not_reported_clients = 0 to UBound(Banked_Month_Client_Array,2)
''		IF Banked_Month_Client_Array(send_to_DHS,not_reported_clients) = False Then
''			objSelection.TypeText "Case # " & Banked_Month_Client_Array(case_num,not_reported_clients) & " for client: " & Banked_Month_Client_Array(clt_first_name, not_reported_clients) & " " & Banked_Month_Client_Array(clt_last_name, not_reported_clients) & " was not added to the DHS Report for Banked Months for the following reason(s):"
''			objSelection.TypeParagraph()
''			objSelection.TypeText "        " & Banked_Month_Client_Array(reason_excluded, not_reported_clients)
''			objSelection.TypeParagraph()
''		End If
''	Next
'End If
'objWord.Visible = True

'Writing to the DHS tracking sheet
'MsgBox "Selct the file of the Excel Spreadsheet you submit to DHS" & _
'	VBNewLine & VBNewLine & "Be sure your spreadsheet is in the correct format."
'Call File_Selection_System_Dialog(list_reported_banked_month_clients)
'call excel_open(list_reported_banked_month_clients, True, True, ObjExcel, objWorkbook)

'For Each objWorkSheet In objWorkbook.Worksheets
'	month_list = month_list & chr(9) & objWorkSheet.Name
'Next


'BeginDialog SNAP_Banked_Month_Report_Dialog, 0, 0, 211, 70, "SNAP Banked Month Reporting Dialog"
'  DropListBox 65, 25, 140, 15, "select one..." & month_list, report_month_dropdown
'  ButtonGroup ButtonPressed
'	OkButton 100, 45, 50, 15
'	CancelButton 155, 45, 50, 15
'  Text 5, 10, 190, 10, "Select the month that you are creating the report for."
'  Text 5, 30, 55, 10, "Month to Report:"
'EndDialog

'Do
'	Dialog SNAP_Banked_Month_Report_Dialog
'	cancel_confirmation
'Loop until report_month_dropdown <> "select one..."

script_end_procedure("Success!")
