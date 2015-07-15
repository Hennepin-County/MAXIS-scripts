'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - ES REFERRAL.vbs"
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
'This function will open the ES_statistics database, check for an existing case and edit it with new info, or add a new entry if there is no existing case in the database.
Function write_MAXIS_info_to_ES_database(ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive, insert_string)
info_array = array(ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive)
	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")


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
			ESMembName = "'" & ESMembName & "'"
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
			ESSite = "'" & ESSite & "'"
		ELSE
			ESSite = "'" & rs("ESSite") & "'"
			IF IsNull(rs("ESSite")) = true THEN ESSite = "null"
		END IF
		IF ESCounselor <> "" THEN 
			ESCounselor = "'" & ESCounselor & "'"
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
						values_string = "'" & item & "'"
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
						values_string = values_string & ", '" & item & "'"
					END IF
				ELSE 
					values_string = values_string & ", null"
				END IF
			END IF
		
		NEXT
		values_string = values_string & ")"
		'Inserting the new record
		objConnection.Execute "INSERT INTO ESTrackingTbl (ESCaseNbr, ESMembNbr, ESMembName, EsSanctionPercentage, ESEmpsStatus, ESTANFMosUsed, ESExtensionReason, ESDisaEnd, ESPrimaryActivity, ESDate, ESSite, ESCounselor, ESActive) VALUES (" & values_string 
		objConnection.Close
	END IF
	'Clearing all variables to avoid writing over records in future calls from same script
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
	
END FUNCTION
			
			

'Collecting ES agencies from database
		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")


		'Opening DB
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & "U:/PHHS/BlueZoneScripts/Statistics/ES statistics.accdb"
		'This looks for an existing case number and edits it if needed
		set rs = objConnection.Execute("SELECT SiteName FROM ESSitesTbl")' Grabbing all ES agency site names
		Dim ES_agencies
		ES_agencies = rs.GetRows()
	objConnection.Close
	set rs = nothing
	
	
	
call convert_array_to_droplist_items (ES_agencies, ES_agency_list)
	
	

Dim case_number, program, referral_date, plan_deadline, ES_provider, other_notes, TIKL_check, dwp_referral_check, es_referral_check, worker_signature, county_collecting_ES_stats
'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog ES_referral_dialog, 0, 0, 276, 235, "Employment services referral"
  ButtonGroup ButtonPressed
    OkButton 145, 210, 50, 15
    CancelButton 205, 210, 50, 15
  EditBox 60, 10, 50, 15, case_number
  DropListBox 215, 10, 50, 10, "DWP"+chr(9)+"MFIP", program
  EditBox 60, 30, 50, 15, referral_date
  EditBox 215, 30, 50, 15, plan_deadline
  EditBox 215, 50, 50, 15, hh_member_list
  DropListBox  85, 90, 185, 15, ES_agency_list, ES_provider
  'EditBox 85, 90, 185, 15, ES_provider
  EditBox 85, 110, 185, 15, other_notes
  CheckBox 10, 70, 230, 15, "Set a TIKL for above date to deny DWP if plan not received.", TIKL_check
  CheckBox 10, 130, 260, 15, "DHS 4161 (DWP referral) sent to client.", dwp_referral_check
  CheckBox 10, 150, 260, 20, "Paper referral sent to ES provider.", es_referral_check
  EditBox 160, 185, 75, 15, worker_signature
  Text 175, 10, 35, 15, "Program:"
  Text 10, 30, 45, 15, "Referral Date:"
  Text 115, 30, 90, 30, "DWP employment plan deadline (10 Business days):"
  Text 55, 50, 155, 15, "HH members referred (separate with commas):"
  Text 10, 90, 55, 15, "ES Provider:"
  Text 10, 110, 70, 15, "Notes:"
  Text 90, 185, 65, 15, "Worker Signature:"
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

'-grabbing case number
EMConnect ""

CALL MAXIS_case_number_finder(case_number)

'-------Calling the dialog / requiring completion of most fields.
DO	
	DO
		DO
			DO 
				DO
					Dialog ES_referral_dialog
					IF ButtonPressed = 0 THEN stopscript
				LOOP UNTIL ButtonPressed = OK
				IF referral_date = "" THEN MsgBox "Please enter a referral date."
			LOOP UNTIL referral_date <> ""
			IF worker_signature = "" THEN MsgBox "Please sign your case note."
		LOOP UNTIL worker_signature <> ""
		IF isnumeric(case_number) = FALSE THEN msgbox "Please enter a valid case number."
	LOOP UNTIL isnumeric(case_number) = True
	IF ES_provider = "" THEN MsgBox "Please enter an employment services provider."
LOOP UNTIL ES_provider <> ""
		

IF TIKL_check = checked THEN
	call navigate_to_MAXIS_screen("DAIL", "WRIT")
	call create_maxis_friendly_date(plan_deadline, 0, 5, 18)
	call write_variable_in_TIKL("If employment plan has not been received, deny DWP.")
	PF3
END IF
'Setting up the HH member list from dialog input
member_array = split(hh_member_list, ",") 
'assigning values for database (necessary to redefine these later so one member doesn't rewrite over existing data)
ESDate = referral_date
ESProvider = ES_Provider


'----Writing the note
call check_for_MAXIS(true)

call start_a_blank_CASE_NOTE

call write_variable_in_CASE_NOTE("* ES REFERRAL SENT *")
call write_variable_in_CASE_NOTE(program & " referral sent to " & ES_provider & " on " & referral_date)
call write_variable_in_CASE_NOTE("Members referred: " & hh_member_list)
IF program = "DWP" THEN call write_bullet_and_variable_in_CASE_NOTE("Employment plan due back on", plan_deadline)
IF dwp_referral_check = checked THEN CALL write_variable_in_CASE_NOTE("DHS 4161 sent to client.")
IF TIKL_check = checked THEN CALL write_variable_in_CASE_NOTE("TIKL to deny DWP this date if employment plan has not been completed.")
IF es_referral_check = checked THEN CALL write_variable_in_CASE_NOTE("Paper referral sent to ES provider.")
IF other_notes <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Other Notes:", other_notes)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)


'Updating the ES Database	(we do this at the end to avoid passing wrong variables to case note
for each member in member_array
	ES_Active = "Yes" 'Assigning new referral values for database
	ES_Counselor = "Unassigned"
	ESPrimary_Activity = "No Employment Plan"
	hh_member = member
	call write_MAXIS_info_to_ES_database(case_number, hh_member, ESMemb_Name, EsSanction_Percentage, ESEmps_Status, ESTANF_MosUsed, ESExtension_Reason, ESDisa_End, ESPrimary_Activity, ESDate, ESprovider, ES_Counselor, ES_active, insert_string)
	ESProvider = ES_provider 'resetting variables for other members
	ESDate = referral_date
next

script_end_procedure("")

	








	

