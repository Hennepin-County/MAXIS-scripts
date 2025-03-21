'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - INTERVIEW TEAM CASE TRACKING.vbs"
start_time = timer
STATS_counter = 0			     'sets the stats counter at one
STATS_manualtime = 	90			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
call changelog_update("01/13/2025", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'TODO  - NEED DIALOG AND REGION TESTING

If NOT IsArray(interviewer_array) Then
	Dim tester_array()
	ReDim tester_array(0)
	Dim interviewer_array()
	ReDim interviewer_array(0)
	tester_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")

	Set fso_command = run_another_script_fso.OpenTextFile(tester_list_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
End If

'Setting the folder names and objects to handle folder and file manipulation
interview_team_cases_folder = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List"
Set objFolder = objFSO.GetFolder(interview_team_cases_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

Const Case_Number_COL 				= 01
Const Application_Date_COL 			= 02
Const Interview_Date_COL 			= 03
Const Interview_Worker_COL 			= 04
Const Interview_Worker_ID_COL		= 05
Const Interview_Length_COL			= 06
Const CASH_COL 						= 07
Const FA_COL 						= 08
Const SNAP_COL 						= 09
Const Expedited_COL 				= 10
Const GRH_COL 						= 11
Const EMER_COL 						= 12
Const EMER_REQ_TYPE_COL 			= 13
Const Addtnl_Intrvwr_Note_COL		= 14
Const CAF_NOTE_Date_COL 			= 15
Const CAF_NOTE_Worker_COL 			= 16
Const Next_Contact_Date_COL			= 17
Const Next_Contact_Worker_COL		= 18
Const Next_Contact_Header_COL		= 19
Const Contacts_Count_COL			= 20
Const NON_CAF_Followup_NOTE_Date_COL = 21
Const NON_CAF_Followup_NOTE_Worker_COL = 22
Const NON_CAF_Followup_NOTE_Header_COL = 23
Const SNAP_APP_COL 					= 24
Const SNAP_Expedited_COL			= 25
Const SNAP_Elig_COL 				= 26
Const CASH_APP_COL 					= 27
Const CASH_Type_COL 				= 28
Const CASH_Elig_COL 				= 29
Const GRH_APP_COL 					= 30
Const GRH_Elig_COL 					= 31
Const EMER_APP_COL 					= 32
Const EMER_Type_COL 				= 33
Const EMER_Elig_COL 				= 34
Const PND2_Denial_Date_COL			= 35
Const Pending_Completed_COL			= 36
Const FILE_NAME_COL 				= 37

' TRACKING NOTES DOC: t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List\Support Documents\Tracking Log Notes and Information.docx"
interview_tracking_excel = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List\Interview Tracking.xlsx"
Call excel_open(interview_tracking_excel, True, False, ObjExcel, objWorkbook)

loged_tracking_records = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List\Loged"

excel_row = 2
all_known_file_names = " "
Do
	file_name = trim(objExcel.Cells(excel_row, FILE_NAME_COL).Value)
	all_known_file_names = all_known_file_names & file_name & " "

	excel_row = excel_row + 1
Loop until file_name = ""
excel_row = excel_row - 1
all_known_file_names = trim(all_known_file_names)
ALL_KNOW_FILES_ARRAY = split(all_known_file_names)

'ADD Cases to Tracking Excel
'creating some objects needed for XML handling
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set xml = CreateObject("Msxml2.DOMDocument")
Dim objTextStream
xml_case_numbs = "*"

'Looking at each xml in the folder for the Interview Team completion
For Each objFile in colFiles								'looping through each file
	file_type = objFile.Type

	If file_type = "XML Source File" Then
		quack = objFile.Name
		file_recorded = False
		xmlPath = objFile.Path												'identifying the current file

		For each duck in ALL_KNOW_FILES_ARRAY
			If duck = quack Then
				file_recorded = True
				With (CreateObject("Scripting.FileSystemObject"))
					.MoveFile xmlPath , loged_tracking_records & "\" & quack & ".xml"
				End With

				Exit For
			End If
		Next

		If file_recorded = False Then
			With (CreateObject("Scripting.FileSystemObject"))
				'Creating an object for the stream of text which we'll use frequently
				If .FileExists(xmlPath) = True then
					name_of_xml = objFile.Name
					If InStr(name_of_xml, "details") Then
						xmlDoc.Async = False

						' Load the XML file
						xmlDoc.load(xmlPath)

						'reads data about the case from the XML
						set node = xmlDoc.SelectSingleNode("//CaseNumber")
						MAXIS_case_number = node.text
						xml_case_numbs = xml_case_numbs & trim(MAXIS_case_number) & "*"

						set node = xmlDoc.SelectSingleNode("//CAFDateStamp")
						app_date = node.text
						app_date = DateAdd("d", 0, app_date)

						set node = xmlDoc.SelectSingleNode("//WorkerName")
						worker_name = node.text

						set node = xmlDoc.SelectSingleNode("//WindowsUserID")
						worker_user_id = node.text

						Set node = xmlDoc.SelectSingleNode("//InterviewLength")
						interview_length = node.text

						set node = xmlDoc.SelectSingleNode("//InterviewDate")
						interview_date = node.text
						interview_date = DateAdd("d", 0, interview_date)

						set node = xmlDoc.SelectSingleNode("//CASHRequest")
						req_cash = node.text
						req_cash = req_cash * 1

						cash_type = ""
						If req_cash = True Then
							set node = xmlDoc.SelectSingleNode("//TypeOfCASH")
							cash_type = node.text
						End If

						If DateDiff("d", #1/13/2025#, interview_date) > 0 Then
							set node = xmlDoc.SelectSingleNode("//GRHRequest")
							req_grh = node.text
							req_grh = req_grh * 1
						End If

						set node = xmlDoc.SelectSingleNode("//SNAPRequest")
						req_snap = node.text
						req_snap = req_snap * 1

						set node = xmlDoc.SelectSingleNode("//EMERRequest")
						req_emer = node.text
						req_emer = req_emer * 1


						set node = xmlDoc.SelectSingleNode("//ExpeditedDetermination")
						exp_det = node.text
						If exp_det <> "" Then exp_det = exp_det * 1

						'Add the file information to the Excel document for the worklist
						' MsgBox "app_date - " & app_date

						ObjExcel.Cells(excel_row, Case_Number_COL).Value 		= MAXIS_case_number
						ObjExcel.Cells(excel_row, Application_Date_COL).Value 	= app_date
						ObjExcel.Cells(excel_row, Interview_Date_COL).Value 	= interview_date
						ObjExcel.Cells(excel_row, Interview_Worker_COL).Value 	= worker_name
						ObjExcel.Cells(excel_row, Interview_Worker_ID_COL).Value= worker_user_id
						ObjExcel.Cells(excel_row, Interview_Length_COL).Value 	= interview_length
						If req_cash = True  Then
							ObjExcel.Cells(excel_row, CASH_COL).Value 									= "TRUE"
							ObjExcel.Cells(excel_row, FA_COL).Value 									= cash_type
						End If
						If req_cash = False Then ObjExcel.Cells(excel_row, CASH_COL).Value 				= "FALSE"
						If req_snap = True  Then
							ObjExcel.Cells(excel_row, SNAP_COL).Value 									= "TRUE"
							If exp_det = True  Then ObjExcel.Cells(excel_row, Expedited_COL).Value		= "TRUE"
							If exp_det = False Then ObjExcel.Cells(excel_row, Expedited_COL).Value 		= "FALSE"
						End If
						If req_snap = False Then ObjExcel.Cells(excel_row, SNAP_COL).Value 				= "FALSE"
						If req_grh  = True  Then ObjExcel.Cells(excel_row, GRH_COL).Value 				= "TRUE"
						If req_grh  = False Then ObjExcel.Cells(excel_row, GRH_COL).Value 				= "FALSE"
						If req_emer = True  Then ObjExcel.Cells(excel_row, EMER_COL).Value 				= "TRUE"
						If req_emer = False Then ObjExcel.Cells(excel_row, EMER_COL).Value 				= "FALSE"

						ObjExcel.Cells(excel_row, FILE_NAME_COL).Value 			= quack

						excel_row = excel_row + 1		'increment the excel row to add more
						STATS_counter = STATS_counter + 1
					End If
				End If
			End With
		End If
	End If
Next

'Looking at each xml in the folder for the Interview Team completion
For Each objFile in colFiles								'looping through each file
	file_type = objFile.Type

	If file_type = "XML Source File" Then
		quack = objFile.Name
		file_recorded = False
		xmlPath = objFile.Path												'identifying the current file

		If file_recorded = False Then
			With (CreateObject("Scripting.FileSystemObject"))
				'Creating an object for the stream of text which we'll use frequently
				If .FileExists(xmlPath) = True then
					name_of_xml = objFile.Name
					If InStr(name_of_xml, "started") Then
						xmlDoc.Async = False

						' Load the XML file
						xmlDoc.load(xmlPath)

						'reads data about the case from the XML
						set node = xmlDoc.SelectSingleNode("//CaseNumber")
						MAXIS_case_number = node.text
						search_numb = "*" & trim(MAXIS_case_number) & "*"

						set node = xmlDoc.SelectSingleNode("//WorkerName")
						worker_name = node.text

						set node = xmlDoc.SelectSingleNode("//WindowsUserID")
						worker_user_id = node.text

						set node = xmlDoc.SelectSingleNode("//ScriptRunDate")
						interview_date = node.text
						interview_date = DateAdd("d", 0, interview_date)

						If InStr(xml_case_numbs, search_numb) = 0 Then

							ObjExcel.Cells(excel_row, Case_Number_COL).Value 		= MAXIS_case_number
							' ObjExcel.Cells(excel_row, Application_Date_COL).Value 	= app_date
							ObjExcel.Range(ObjExcel.Cells(excel_row, Application_Date_COL), ObjExcel.Cells(excel_row, Application_Date_COL)).Interior.ColorIndex = 3			'RED
							ObjExcel.Cells(excel_row, Interview_Date_COL).Value 	= interview_date
							ObjExcel.Cells(excel_row, Interview_Worker_COL).Value 	= worker_name
							ObjExcel.Cells(excel_row, Interview_Worker_ID_COL).Value= worker_user_id
							ObjExcel.Range(ObjExcel.Cells(excel_row, Interview_Length_COL), ObjExcel.Cells(excel_row, Interview_Length_COL)).Interior.ColorIndex = 3			'RED

							ObjExcel.Cells(excel_row, FILE_NAME_COL).Value 			= quack

							excel_row = excel_row + 1		'increment the excel row to add more
							STATS_counter = STATS_counter + 1
						End If
					End If
				End If
			End With
		End If
	End If
Next
objWorkbook.Save()		'saving the excel

transmit
Call check_for_MAXIS(False)

'READ for Follow Up Notes
excel_rows_with_processing_completed = " "
excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, Case_Number_COL).Value)

	Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Pending_Completed_COL), is_case_completed)
	interviewer_name =  trim(ObjExcel.Cells(excel_row, Interview_Worker_COL))
	If is_case_completed <> True or interviewer_name = "" Then
		STATS_counter = STATS_counter + 1

		app_date = trim(ObjExcel.Cells(excel_row, Application_Date_COL).Value)
		If app_date <> "" Then 			'For some tracking files we do not have all the information and so we cannot do all the things
			app_date = DateAdd("d", 0, app_date)
			Call convert_date_into_MAXIS_footer_month(app_date, MAXIS_footer_month, MAXIS_footer_year)

			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
			application_processed = False
			If case_pending = False Then application_processed = True
			If case_pending = True Then
				If unknown_cash_pending = False and ga_status <> "PENDING" and msa_status <> "PENDING" and mfip_status <> "PENDING" and dwp_status <> "PENDING" and grh_status <> "PENDING" and snap_status <> "PENDING" and emer_status <> "PENDING" Then application_processed = True
			End If
			If application_processed = True Then excel_rows_with_processing_completed = excel_rows_with_processing_completed & excel_row & " "

			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, EMER_COL), emer_request)
			emer_pend_type = trim(ObjExcel.Cells(excel_row, EMER_REQ_TYPE_COL))
			' MsgBox "emer_request - " & emer_request & vbCr & "emer_pend_type - " & emer_pend_type
			If emer_request = True and emer_pend_type = "" Then
				' MsgBox "IN IT" & vbCr & "emer_type - " & emer_type
				ObjExcel.Cells(excel_row, EMER_REQ_TYPE_COL) = emer_type
			End If


			interview_date = trim(ObjExcel.Cells(excel_row, Interview_Date_COL).Value)
			interview_date = DateAdd("d", 0, interview_date)
			CAF_Note_date = trim(ObjExcel.Cells(excel_row, CAF_NOTE_Date_COL))
			FollowUp_Note_date = trim(ObjExcel.Cells(excel_row, NON_CAF_Followup_NOTE_Date_COL))

			CAF_date = ""
			CAF_worker = ""
			FollowUp_date = ""
			FollowUp_worker = ""
			FollowUp_text = ""
			Contact_date = ""
			Contact_worker = ""
			Contact_text = ""
			Contact_count = 0
			PND2_denial_date = ""
			additional_interviewer_note = False

			interviewer_name =  trim(ObjExcel.Cells(excel_row, Interview_Worker_COL))
			interviewer_user_id = trim(ObjExcel.Cells(excel_row, Interview_Worker_ID_COL))
			interviewer_mx_number = ""
			run_by_interview_team = False										'Default the interview team option to false
			If trim(interviewer_name) = "" and trim(interviewer_user_id) <> "" Then
				For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
					If UCase(interviewer_user_id) = UCase(worker.interviewer_id_number) Then
						interviewer_name = worker.interviewer_full_name
						interviewer_mx_number = UCase(worker.interviewer_x_number)
						ObjExcel.Cells(excel_row, Interview_Worker_COL) = worker.interviewer_full_name
						Exit For
					End If
				Next
			Else
				For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
					If UCase(interviewer_name) = UCase(worker.interviewer_full_name) Then
						interviewer_mx_number = UCase(worker.interviewer_x_number)
						Exit For
					End If
				Next
			End If

			If CAF_Note_date = "" or (application_processed = True and FollowUp_Note_date = "") Then

				Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)               'Now we navigate to CASE:NOTES
				If is_this_priv = False Then
					too_old_date = DateAdd("D", -1, interview_date)              'We don't need to read notes from before the CAF date

					note_row = 5
					Do
						EMReadScreen note_date, 	8, note_row, 6                  'reading the note date
						EMReadScreen note_worker, 	7, note_row, 16
						EMReadScreen note_title, 	55, note_row, 25               'reading the note header
						note_title = trim(note_title)

						If left(note_title, 24) = "~ Interview Completed on" Then
							If interviewer_name = "" Then
								For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
									If UCase(note_worker) = UCase(worker.interviewer_x_number) Then
										interviewer_name = worker.interviewer_full_name
										ObjExcel.Cells(excel_row, Interview_Worker_COL).Value = interviewer_name
										Exit For
									End If
								Next
								Exit Do
							End If
						End If
						' If note_title = "Processing Needed: follow up notes from Interview" Then Exit Do
						If note_date = "        " Then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

						exclude_this_note = False
						If note_title = "~~~continued from previous note~~~" Then exclude_this_note = True
						If left(note_title, 11) = "APPROVAL - " Then exclude_this_note = True
						If left(note_title, 16) =  "REVW COMPLETE - " Then exclude_this_note = True
						If left(note_title, 16) = "MFIP Orientation" Then exclude_this_note = True
						If note_title = "**AREP removed**" Then exclude_this_note = True
						If note_title = "Phone Interview Attempted but Interview NOT Completed" Then exclude_this_note = True
						If interviewer_mx_number = note_worker Then exclude_this_note = True		'we are not interested in notes by the interviewing worker.
						If note_worker = "X127L1S" Then exclude_this_note = True
						If note_worker = "X127DC6" Then exclude_this_note = True
						If note_worker = "X127MR7" Then exclude_this_note = True
						If left(note_title, 24) = "~ Application Received (" Then exclude_this_note = True
						If left(note_title, 23) = "~ HC Pended from a METS" Then exclude_this_note = True
						If left(note_title, 32) = "~ Received Application for SNAP," Then exclude_this_note = True
						If left(note_title, 33) = "~ Appointment letter sent in MEMO" Then exclude_this_note = True
						If left(note_title, 34) = "Subsequent Application Requesting:" Then exclude_this_note = True
						If InStr(note_title, " EBT TRX ") <> 0 Then exclude_this_note = True
						If InStr(note_title, "MONY/CHCK ISSUED ON ") <> 0 Then exclude_this_note = True
						If InStr(note_title, " PARIS Match (") <> 0 Then exclude_this_note = True
						If InStr(note_title, "Paris match -  ") <> 0 Then exclude_this_note = True
						If InStr(note_title, "WAGE MATCH (") <> 0 Then exclude_this_note = True
						If InStr(note_title, "WAGE MATCH(") <> 0 Then exclude_this_note = True
						If InStr(note_title, "-----Appeal ") <> 0 Then exclude_this_note = True
						If InStr(note_title, "DISB ") <> 0 and note_worker = "CS DAIL" Then exclude_this_note = True
						If left(note_worker, 2) = "PW" Then exclude_this_note = True
						If left(note_title, 25) = "Client picked up EBT card" Then exclude_this_note = True
						If note_title = "EDRS RAN FOR THE CASE" Then exclude_this_note = True
						If left(note_title, 22) = "DD set on this account" Then exclude_this_note = True
						If left(note_title, 13) = "*** EX PARTE " Then exclude_this_note = True
						If note_title = "~ Client has not completed application interview, NOMI" Then exclude_this_note = True
						If InStr(note_title, "HC Certain Pops App:") Then exclude_this_note = True
						If InStr(note_title, "HC Renewal Form:") Then exclude_this_note = True
						If InStr(note_title, "Ex Parte Renewal") Then exclude_this_note = True
						If InStr(note_title, "DHS-3876 (Certain Pops App)") Then exclude_this_note = True
						If InStr(UCASE(note_title), "OOPS") Then exclude_this_note = True
						If left(note_title, 7) = "**EBT**" Then exclude_this_note = True

						' MsgBox "interviewer_mx_number - " & interviewer_mx_number & vbCr & "note_worker - " & note_worker

						If Instr(note_title, " HUF for ") <> 0 or Instr(note_title, " CAF for ") <> 0 Then
							CAF_date = note_date
							CAF_worker = note_worker
						ElseIf Instr(note_title, "Office visit at") <> 0 or Instr(note_title, "Office visit from") <> 0 or InStr(note_title, "Phone call from") <> 0 or InStr(note_title, "Voicemail to") <> 0 or InStr(note_title, "Voicemail from") <> 0  or InStr(note_title, "Chat Message from") <> 0 or InStr(note_title, "Infokeep chat from") <> 0 or InStr(note_title, "Voicemail assignment") Then		'or InStr(note_title, "Phone call to") <> 0
							Contact_date = note_date
							Contact_worker = note_worker
							Contact_text = note_title
							Contact_count = Contact_count + 1
						ElseIf InStr(note_title, "DENIAL of ") <> 0 Then
							PND2_denial_date = note_date
						ElseIF exclude_this_note = False Then
							additional_contact_with_interview_team = False
							For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
								' MsgBox "NOTE WORKER - " & note_worker & vbCr & vbCr & "worker x number - " & worker.interviewer_x_number & vbCr & "trainer - " & worker.interview_trainer
								If worker.interview_trainer = False and UCase(worker.interviewer_x_number) <> "X127XXX" and worker.interviewer_x_number <> "" Then
									If note_worker = UCase(worker.interviewer_x_number) Then
										' MsgBox "MATCH !!!" & vbCr & vbCr & "NOTE WORKER - " & note_worker & vbCr & vbCr & "worker x number - " & worker.interviewer_x_number & vbCr & "trainer - " & worker.interview_trainer
										additional_contact_with_interview_team = True
										Contact_date = note_date
										Contact_worker = note_worker
										Contact_text = note_title
										Contact_count = Contact_count + 1
										Exit Do
									End If
								End If
							Next
							If additional_contact_with_interview_team = False Then
								FollowUp_date = note_date
								FollowUp_worker = note_worker
								FollowUp_text = note_title
							End If
						ElseIf interviewer_mx_number = note_worker Then
							' MsgBox "MATCH !!!" & vbCr & vbCr & "interviewer_mx_number - " & interviewer_mx_number & vbCr & "note_worker - " & note_worker & vbCr & vbCr & "note_title - " & note_title

							If note_title <> "Processing Needed: follow up notes from Interview" and note_title <> "~~~continued from previous note~~~" and left(note_title, 24) <> "~ Interview Completed on" and note_title <> "VERIFICATIONS REQUESTED" and left(note_title, 25) <> "Expedited Determination: " and note_title <> "INTERVIEW INCOMPLETE - Attempt made but additional deta" Then
								' MsgBox "SAVE !!!" & vbCr & vbCr & "interviewer_mx_number - " & interviewer_mx_number & vbCr & "note_worker - " & note_worker & vbCr & vbCr & "note_title - " & note_title
								additional_interviewer_note = True
							End If
						End If


						note_row = note_row + 1
						If note_row = 19 Then
							note_row = 5
							PF8
							EMReadScreen check_for_last_page, 9, 24, 14
							If check_for_last_page = "LAST PAGE" Then Exit Do
						End If
						EMReadScreen next_note_date, 8, note_row, 6
						If next_note_date = "        " Then Exit Do
					Loop until DateDiff("d", too_old_date, next_note_date) <= 0
				Else
					FollowUp_text = "PRIVILEGED CASE"
				End If
			End If

			If CAF_date <> "" Then
				ObjExcel.Cells(excel_row, CAF_NOTE_Date_COL).Value = CAF_date
				ObjExcel.Cells(excel_row, CAF_NOTE_Worker_COL).Value = CAF_worker
			End If
			If FollowUp_date <> "" Then
				' MsgBox "excel_row - " & excel_row & vbCr & "FollowUp_text - " & FollowUp_text
				If left(FollowUp_text, 1) = "=" Then FollowUp_text = "'" & FollowUp_text
				ObjExcel.Cells(excel_row, NON_CAF_Followup_NOTE_Date_COL).Value = FollowUp_date
				ObjExcel.Cells(excel_row, NON_CAF_Followup_NOTE_Worker_COL).Value = FollowUp_worker
				ObjExcel.Cells(excel_row, NON_CAF_Followup_NOTE_Header_COL).Value = FollowUp_text
			End If
			If Contact_date <> "" Then
				If left(Contact_text, 1) = "=" Then Contact_text = "'" & Contact_text
				ObjExcel.Cells(excel_row, Next_Contact_Date_COL).Value 		= Contact_date
				ObjExcel.Cells(excel_row, Next_Contact_Worker_COL).Value 	= Contact_worker
				ObjExcel.Cells(excel_row, Next_Contact_Header_COL).Value 	= Contact_text
				ObjExcel.Cells(excel_row, Contacts_Count_COL).Value 	= Contact_count
			End If
			ObjExcel.Cells(excel_row, PND2_Denial_Date_COL).Value = PND2_denial_date
			If additional_interviewer_note = True Then ObjExcel.Cells(excel_row, Addtnl_Intrvwr_Note_COL).Value = "TRUE"
		End If
	End If
	Call back_to_SELF
	excel_row = excel_row + 1
	' If excel_row = 21 Then Exit Do		'TESTING CODE
	next_case = trim(ObjExcel.Cells(excel_row, Case_Number_COL).Value)
Loop until next_case = ""
objWorkbook.Save()		'saving the excel

transmit
Call check_for_MAXIS(False)

'READ for approvals
excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, Case_Number_COL).Value)

	Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Pending_Completed_COL), is_case_completed)
	If is_case_completed <> True Then
		STATS_counter = STATS_counter + 1

		app_date = trim(ObjExcel.Cells(excel_row, Application_Date_COL).Value)
		If app_date <> "" Then 			'For some tracking files we do not have all the information and so we cannot do all the things
			app_date = DateAdd("d", 0, app_date)
			interview_date = trim(ObjExcel.Cells(excel_row, Interview_Date_COL).Value)
			interview_date = DateAdd("d", 0, interview_date)

			Call convert_date_into_MAXIS_footer_month(app_date, MAXIS_footer_month, MAXIS_footer_year)

			Call navigate_to_MAXIS_screen("ELIG", "SUMM")
			EMReadScreen numb_DWP_versions, 		1, 7, 40
			EMReadScreen numb_MFIP_versions, 		1, 8, 40
			EMReadScreen numb_MSA_versions, 		1, 11, 40
			EMReadScreen numb_GA_versions, 			1, 12, 40
			EMReadScreen numb_CASH_denial_versions, 1, 13, 40
			EMReadScreen numb_GRH_versions, 		1, 14, 40
			EMReadScreen numb_EMER_versions, 		1, 16, 40
			EMReadScreen numb_SNAP_versions, 		1, 17, 40

			If numb_SNAP_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "FS  ")
				EMReadScreen on_elig_fs, 4, 3, 48
				If on_elig_fs = "FSPR" Then
					Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, SNAP_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, SNAP_Elig_COL).Value = elig_version_result

							If elig_version_result = "ELIGIBLE" Then
								snap_expedited = False
								transmit 		'FSCR
								EMReadScreen case_expedited_indicator, 9, 4, 3
								If case_expedited_indicator = "EXPEDITED" Then snap_expedited = True
								If snap_expedited = True Then ObjExcel.Cells(excel_row, SNAP_Expedited_COL).Value = "TRUE"
								If snap_expedited = False Then ObjExcel.Cells(excel_row, SNAP_Expedited_COL).Value = "FALSE"
							End If

						End If
					End If
				End If
				PF3
			End If

			If numb_EMER_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "EMER")
				Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
				If approved_version_found = True Then
					If DATEDiff("d", app_date, elig_version_date) > 0 Then
						EMReadScreen the_type, 3, 4, 45
						ObjExcel.Cells(excel_row, EMER_Type_COL).Value = trim(the_type)
						ObjExcel.Cells(excel_row, EMER_APP_COL).Value = elig_version_date
						ObjExcel.Cells(excel_row, EMER_Elig_COL).Value = elig_version_result
					End If
				End If
				PF3
			End If

			If numb_GRH_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "GRH ")
				EMReadScreen on_elig_grh, 4, 3, 47
				If on_elig_grh = "GRPR" Then
					Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, GRH_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, GRH_Elig_COL).Value = elig_version_result
						End If
					End If
				End If
				PF3
			End If

			If numb_CASH_denial_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "DENY")
				EMReadScreen on_elig_deny, 4, 3, 48
				If on_elig_deny = "CAPR" Then
					Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, CASH_Type_COL).Value = "None"
							ObjExcel.Cells(excel_row, CASH_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, CASH_Elig_COL).Value = elig_version_result
						End If
					End If
				End If
				PF3
			End If

			If numb_DWP_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "DWP ")
				EMReadScreen on_elig_dwp, 4, 3, 47
				If on_elig_dwp = "DWPR" Then
					Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, CASH_Type_COL).Value = "DWP"
							ObjExcel.Cells(excel_row, CASH_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, CASH_Elig_COL).Value = elig_version_result
						End If
					End If
				End If
				PF3
			End If

			If numb_MSA_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "MSA ")
				EMReadScreen on_elig_msa, 4, 3, 47
				If on_elig_msa = "MSPR" Then
					Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, CASH_Type_COL).Value = "MSA"
							ObjExcel.Cells(excel_row, CASH_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, CASH_Elig_COL).Value = elig_version_result
						End If
					End If
				End If
				PF3
			End If

			If numb_GA_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "GA  ")
				EMReadScreen on_elig_ga, 4, 3, 48
				If on_elig_ga = "GAPR" Then
					Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, CASH_Type_COL).Value = "GA"
							ObjExcel.Cells(excel_row, CASH_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, CASH_Elig_COL).Value = elig_version_result
						End If
					End If
				End If
				PF3
			End If

			If numb_MFIP_versions <> " " Then
				call navigate_to_MAXIS_screen("ELIG", "MFIP")
				EMReadScreen on_elig_mfip, 4, 3, 47
				If on_elig_mfip = "MFPR" Then
					Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
					If approved_version_found = True Then
						If DATEDiff("d", app_date, elig_version_date) > 0 Then
							ObjExcel.Cells(excel_row, CASH_Type_COL).Value = "MFIP"
							ObjExcel.Cells(excel_row, CASH_APP_COL).Value = elig_version_date
							ObjExcel.Cells(excel_row, CASH_Elig_COL).Value = elig_version_result
						End If
					End If
				End If
				PF3
			End If
		End If
	End If
	Call back_to_SELF
	excel_row = excel_row + 1
	' If excel_row = 21 Then Exit Do		'TESTING CODE

	next_case = trim(ObjExcel.Cells(excel_row, Case_Number_COL).Value)
Loop until next_case = ""
objWorkbook.Save()		'saving the excel

excel_rows_with_processing_completed = trim(excel_rows_with_processing_completed)
complete_these_rows_array = split(excel_rows_with_processing_completed)
For each excel_row in complete_these_rows_array
	excel_row = excel_row * 1
	ObjExcel.Cells(excel_row, Pending_Completed_COL) = "TRUE"
Next
objWorkbook.Save()		'saving the excel

call script_end_procedure("All Done")