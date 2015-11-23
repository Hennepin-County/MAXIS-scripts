'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - ES DB MAINTENANCE.vbs"
start_time = timer

'Option Explicit

DIM beta_agency
DIM FuncLib_URL, req, fso

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'Declaring Variables
Dim agency, update_type, received_sent, case_number, document_date, hh_member, program_list, ES_provider, ES_counselor, primary_activity, activity_hours, FSS_check, UP_check, other_check, job_info, job_verif_check, school_info, school_verif_check, disa_end_date, mof_check, actions_taken, other_notes, worker_signature

BeginDialog UPDATE_ES_DB_dialog, 0, 0, 218, 120, "ES Database Update Dialog"
  EditBox 84, 20, 130, 15, worker_number
  CheckBox 4, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 109, 100, 50, 15
    CancelButton 164, 100, 50, 15
  Text 4, 25, 65, 10, "Worker(s) to check:"
  Text 4, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 14, 5, 125, 10, "***UPDATE FAMILY CASH EMPLOYMENT SERVICES DATABASE***"
  Text 4, 40, 210, 20, "Enter last 3 digits of your workers' x1 numbers (ex: x100###), separated by a comma."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------
'THIS SCRIPT WILL POPULATE THE ES RECORDS ACCESS DATABASE WITH INFORMATION FROM REPT/MFCM, IT THEN WILL SCAN BACK THROUGH ANY
'RECORDS WITHOUT AN ES PROVIDER AND CHECK INFC/WORK FOR MOST RECENT REFERRAL.  CURRENTLY HAS SAINT LOUIS COUNTY PROVIDER NUMBERS 
'HARD CODED, THESE SHOULD BE MOVED TO THE DB AND READ FROM THERE FOR FUTURE COMPATIBILITY WITH OTHER COUNTIES.

'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog UPDATE_ES_DB_dialog
If buttonpressed = cancel then stopscript

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
PF3
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You appear to be locked out of MAXIS.")

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = worker_county_code & trim(replace(ucase(x1_number), worker_county_code, ""))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & worker_county_code & trim(replace(ucase(x1_number), worker_county_code, "")) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

	

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_screen("rept", "mfcm")
	EMWriteScreen worker, 21, 13
	transmit
	
	'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
  has_content_check = trim(has_content_check)
	If has_content_check <> "" then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/MFCF it displays right away, instead of when the second F8 is sent
			Do			
				EMReadScreen case_number, 8, MAXIS_row, 6		  'Reading case number  ADD LOGIC FOR IF BLANK, USE PREVIOUS
				EMReadScreen client_name, 20, MAXIS_row, 16		'Reading client name
				EMReadScreen sanc_perc, 2, MAXIS_row, 39	    'Reading Sanction Percentage
				EMReadScreen vend_rsn, 2, MAXIS_row, 45		    'Reading Vend Rsn
				EMReadScreen emps_status, 2, MAXIS_row, 52		'Reading Emps Status
				EMReadScreen hrs_retro, 3, MAXIS_row, 57			'Reading Hrs Retro
				EMReadScreen empl_pro, 3, MAXIS_row, 62			  'Reading Empl Pro
				EMReadScreen tanf_mos, 2, MAXIS_row, 69			  'Reading TANF Mos
				EMReadScreen sixty_ext_rsn, 2, MAXIS_row, 75	'Reading 60 Mos Ext Rsn

				'Doing this because  sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " and client_name = "                    " then exit do			'Exits do if we reach the end
				
				'Needs to go into MAXIS to get member number
				IF case_number = "        " THEN 'This handles the second member on a household, which won't have the case number visible.
					case_number = previous_case_number
					EMWriteScreen "S", (MAXIS_row - 1), 3
				ELSE 'Normal handling
					EMWriteScreen "S", MAXIS_row, 3
				END IF
				Transmit
				EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
				IF priv_check <> "PRIVIL" THEN 'Non priv case, read everything and update Database
					EMWriteScreen "MEMB", 20, 71
					Transmit
					row = 1
					col = 1
					'the following mess is formatting the name to match STAT formatting
					truncated_client_name_array = split(replace(client_name, ",", ""))
					DO
						IF len(truncated_client_name_array(0)) < 5 THEN truncated_client_name_array(0) = truncated_client_name_array(0) & " "
					LOOP UNTIL len(truncated_client_name_array(0)) >= 5
					DO
						IF len(truncated_client_name_array(1)) < 7 THEN truncated_client_name_array(1) = truncated_client_name_array(1) & " "
					LOOP UNTIL len(truncated_client_name_array(1)) >= 7
					truncated_client_name = left(truncated_client_name_array(0), 5) & " " & left(truncated_client_name_array(1), 7) & " " & truncated_client_name_array(2)
					'msgbox truncated_client_name
					EMSearch truncated_client_name, row, col 'so it can find the member number here
					EMReadScreen hh_member, 2, row, 3
					if hh_member = "AR" then hh_member = 124 'This handles the error if it can't find the member name.  Later we can query the database for all member 124's
					PF3
					'msgbox case_number & hh_member
					ESActive = "Yes"
					IF emps_status = 10 THEN ESActive = "No"
				'Getting rid of spaces from variables for database
					case_number = replace(case_number, " ", "")
					'client_name = replace(client_name, " ", "")
					sanc_perc = replace(sanc_perc, " ", "")
					vend_rsn = replace(vend_rsn, " ", "")
					emps_status = replace(emps_status, " ", "")
					hrs_retro = replace(hrs_retro, " ", "")
					empl_pro = replace(empl_pro, " ", "")
					tanf_mos = replace(tanf_mos, " ", "")
					sixty_ext_rsn = replace(sixty_ext_rsn, " ", "")
				
				'Creating object for access
					Set objConnection = CreateObject("ADODB.Connection")
					Set objRecordSet = CreateObject("ADODB.Recordset")
					
				'Put the variables into an array for syntax conversion
					info_array = array(case_number, hh_member, client_name, sanc_perc, emps_status, tanf_mos, sixty_ext_rsn, ESActive)
					'Opening DB
					objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "U:/PHHS/BlueZoneScripts/Statistics/ES statistics.accdb"
						'This looks for an existing case number and edits it if needed
					set rs = objConnection.Execute("SELECT * FROM ESTrackingTbl WHERE ESCaseNbr = " & case_number & " AND ESMembNbr = " & hh_member & "") 'pulling all existing case / member info into a recordset
					
							
					IF NOT(rs.EOF) THEN 'There is an existing case, we need to update
						'we don't want to overwrite existing data that isn't updated by the script, 
						'the following IF/THENs assign variables to the value from the recordset/database for variables that are empty in the script, and if already null in database,
						'set to "null" for inclusion in sql string.  Also appending quotes / hashtags for string / date variables.
						IF ESCaseNbr = "" THEN ESCaseNbr = rs("ESCaseNbr") 'no null setting, should never happen, but just in case we do not want to ever overwrite a case number / member number
						IF ESMembNbr = "" THEN ESMembNbr = rs("ESMembNbr")
						IF ESMembName <> "" THEN 
							ESMembName = "'" & replace(ESMembName, "'", "") & "'"
						ELSE
							ESMembName = "'" & rs("ESMembName") & "'"
							IF IsNull(rs("ESMembName")) = true THEN ESMembName = "null"
						END IF
						IF ESSanctionPercentage = "" THEN
							ESSanctionPercentage = rs("ESSanctionPercentage")
							IF IsNull(rs("ESSanctionPercentage")) = true THEN ESSanctionPercentage = "null"
						END IF
						IF ESEmpsStatus = "" THEN 
							ESEmpsStatus = rs("ESEmpsStatus")
							IF IsNull(rs("ESEmpsStatus")) = true THEN ESEmpsStatus = "null"
						END IF
						IF ESTANFMosUsed = "" THEN
							ESTANFMosUsed = rs("ESTANFMosUsed")
							IF ISNull(rs("ESTANFMosUsed")) = true THEN ESTANFMosUsed = "null"
						END IF
						IF ESExtensionReason = "" THEN 
							ESExtensionReason = rs("ESExtensionReason")
							IF IsNull(rs("ESExtensionReason")) = true THEN ESExtensionReason = "null"
						END IF
						IF IsDate(ESDisaEnd) = TRUE THEN 
							ESDisaEnd = "#" & ESDisaEnd & "#"
						ELSE
							IF ESDisaEnd = "" THEN ESDisaEnd = "#" & rs("ESDisaEnd") & "#"
							IF IsNull(rs("ESDisaEnd")) = true THEN ESDisaEnd = "null"
						END IF
						IF ESPrimaryActivity <> "" THEN 
							ESPrimaryActivity = "'" & ESPrimaryActivity & "'"
						ELSE
							ESPrimaryActivity = "'" & rs("ESPrimaryActivity") & "'"
							IF IsNull(rs("ESPrimaryActivity")) = true THEN ESPrimaryActivity = "null"
						END IF
						IF IsDate(ESDate) = True THEN
							ESDate = "#" & ESDate & "#"
						ELSE
							ESDate = "#" & rs("ESDate") & "#"
							IF IsNull(rs("ESDate")) = true THEN ESDate = "null"
						END IF
						IF ESSite <> "" THEN 
							ESSite = "'" & replace(ESSite, "'", "") & "'"
						ELSE
							ESSite = "'" & rs("ESSite") & "'"
							IF IsNull(rs("ESSite")) = true THEN ESSite = "null"
						END IF
						IF ESCounselor <> "" THEN 
							ESCounselor = "'" & replace(ESCounselor, "'", "") & "'"
						ELSE
							ESCounselor = "'" & rs("ESCounselor") & "'"
							IF IsNull(rs("ESCounselor")) = true THEN ESCounselor = "null"
						END IF
						IF ESActive <> "" THEN 
							ESActive = "'" & ESActive & "'"
						ELSE
							ESActive = "'" & rs("ESActive") & "'"
							IF IsNull(rs("ESActive")) = true THEN ESActive = "null"
						END IF
						'This formats all the variables into the correct syntax 	
						ES_update_str = "ESMembName = " & ESMembName & ", ESSanctionPercentage = " & ESSanctionPercentage & ", ESEmpsStatus = " & ESEmpsStatus & ", ESTANFMosUsed = " & ESTANFMosUsed &_
								", ESExtensionReason = " & ESExtensionReason & ", ESDisaEnd = " & ESDisaEnd & ", ESPrimaryActivity = " & ESPrimaryActivity & ", ESDate = " & ESDate & ", ESSite = " &_
								ESSite & ", ESCounselor = " & ESCounselor & ", ESActive = " & ESActive & " WHERE ESCaseNbr = " & ESCaseNbr & " AND ESMembNbr = " & ESMembNbr & ""
						objConnection.Execute "UPDATE ESTrackingTbl SET " & ES_update_str 'Here we are actually writing to the database
						'msgbox ES_update_str
						objConnection.Close 
						set rs = nothing
					ELSE 'There is no existing case, add a new one using the info pulled from the script
						FOR EACH item IN info_array ' THIS loop writes the values string for the SQL statement (with correct syntax for each variable type) to write a NEW RECORD to the database
							IF values_string = "" THEN 
								IF item <> "" THEN 
									IF isnumeric(item) = true THEN
										values_string = """ " & item & " """
									ELSEIF isdate(item) = true Then
										values_string = " #" & item & "#"
									ELSE
										values_string = "'" & replace(item, "'", "") & "'"
									END IF
								ELSE 
									values_string = "null"
								END IF
							ELSE
								IF item <> "" THEN
									IF isnumeric(item) = true THEN
										values_string = values_string & ", "" " & item & " """
									ELSEIF isdate(item) = true THEN
										values_string = values_string & ", #" & item & "#"
									ELSE
										values_string = values_string & ", '" & replace(item, "'", "") & "'"
									END IF
								ELSE 
									values_string = values_string & ", null"
								END IF
							END IF
						
						NEXT
						values_string = values_string & ")"
						'msgbox values_string
						'Inserting the new record
						objConnection.Execute "INSERT INTO ESTrackingTbl (ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESActive) VALUES (" & values_string 
						objConnection.Close
					END IF
					'Clearing all variables to avoid writing over records 
					ERASE info_array
					ESMembNbr = "" 
					ESMembName = "" 
					EsSanctionPercentage = "" 
					ESEmpsStatus = "" 
					ESTANFMosUsed = "" 
					ESExtensionReason = "" 
					ESDisaEnd = "" 
					ESPrimaryActivity = "" 
					ESDate = "" 
					ESSite = "" 
					ESCounselor = ""
					ESActive = ""
					insert_string = ""
					values_string = ""
				ELSE 'This is for priv cases - we can't pull a member number, so don't want to update incorrectly, simply saves case number
					prived_case_list = prived_case_list & case_number & ", "
					'msgbox prived_case_list
				END IF		
				'Finding the new row (the screen changes when we come back from STAT, so it searches for the name it just processed and adds 1 row.
				row = 7
				col = 1
				EMSearch client_name, row, col
				IF row = 0 THEN row = 6 'Sets back to 1st row if can't find the name to make sure no one is missed
				MAXIS_row = row + 1
				previous_case_number = case_number 'saving the case number for the second household member
				case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

msgbox "MFCM data written to database, the script will now attempt to update ES providers."

'Creating object for access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
					
objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "U:/PHHS/BlueZoneScripts/Statistics/ES statistics.accdb"
	'Pulling a recordset of all records without an ESSite entered
set rs = objConnection.Execute("SELECT * FROM ESTrackingTbl WHERE ESSite IS NULL")
'objConnection.Close		'closing the connection while we pull the data out of MAXIS			

IF NOT(rs.eof) THEN 'Do the following for any existing records
	msgbox "It does exist!"
	rs.MoveFirst
	DO 
		case_number = rs("ESCaseNbr")
		call navigate_to_MAXIS_screen("INFC", "WORK")
		EMReadScreen no_referral_check, 6, 24, 2
		IF no_referral_check <> "NO WF1" THEN
			member_number = rs("ESMembNbr")
			IF len(member_number) = 1 THEN member_number = "0" & member_number
			row = 1
			col = 1
			EMSearch member_number, row, col
			IF row = 0 THEN provider = rs("ESSite")
			IF Isnull(provider) = True THEN provider = "null"
			EMReadScreen provider_number, 6, row, 49
			IF provider_number = "000046" THEN provider = "'NEMOJT - VIRGINIA'"
			IF provider_number = "000137" THEN provider = "'AEOA - DULUTH'"
			IF provider_number = "000048" THEN provider = "'NEMOJT - DULUTH'"
			IF provider_number = "000047" THEN provider = "'NEMOJT - HIBBING'"
			IF provider_number = "000131" THEN provider = "'AEOA - VIRGINIA'"
			IF provider_number = "000133" THEN provider = "'AEOA - HIBBING'"
			IF provider_number = "000152" THEN provider = "'DWD'"
			IF provider_number = "000249" THEN provider = "'MCT - VIRGINIA'"
			IF provider_number = "000250" THEN provider = "'MCT - DULUTH'"
			IF provider_number = "000297" THEN provider = "'CAD'"
			
			 'writing the info to the DB
			' objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "U:/PHHS/BlueZoneScripts/Statistics/ES statistics.accdb"
			 objConnection.Execute "UPDATE ESTrackingTbl SET ESSite = " & provider & " WHERE ESCaseNbr = " & case_number & " AND ESMembNbr = " & member_number & ""
			 'objconnection.Close
		END IF	
		rs.MoveNext
	LOOP UNTIL(rs.eof = true)
END IF
objConnection.close
set rs = nothing

script_end_procedure




