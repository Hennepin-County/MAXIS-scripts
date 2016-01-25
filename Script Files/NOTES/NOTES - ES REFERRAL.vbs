'Option Explicit
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - ES REFERRAL.vbs"
start_time = timer

'DIM name_of_script, start_time, FuncLib_URL, run_locally, default_directory, beta_agency, req, fso

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

IF collecting_ES_statistics = true THEN
'Collecting ES agencies from database
		'Setting constants
		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")


		'Opening DB
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & ES_database_path
		'This looks for an existing case number and edits it if needed
		set rs = objConnection.Execute("SELECT SiteName FROM ESSitesTbl")' Grabbing all ES agency site names
		Dim ES_agencies
		ES_agencies = rs.GetRows() 'This returns the contents of the recordset as an array
	objConnection.Close
	set rs = nothing

	call convert_array_to_droplist_items (ES_agencies, ES_agency_list) 'This converts the array of ES Agencies into a droplist for dialog
END IF
	

'Dim case_number, program, referral_date, plan_deadline, ES_provider, other_notes, TIKL_check, dwp_referral_check, es_referral_check, worker_signature, county_collecting_ES_stats
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
  IF collecting_ES_statistics = True THEN
	DropListBox  85, 90, 185, 15, ES_agency_list, ES_provider
  ELSE
	EditBox 85, 90, 185, 15, ES_provider
  END IF
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
	err_msg = ""
	Dialog ES_referral_dialog
	IF ButtonPressed = 0 THEN stopscript
	IF referral_date = "" THEN err_msg = err_msg & vbCr & "Please enter a referral date."
	IF worker_signature = "" THEN err_msg = err_msg & vbCr &  "Please sign your case note."
	IF isnumeric(case_number) = FALSE THEN err_msg = err_msg & vbCr &  "Please enter a valid case number."
	IF ES_provider = "" THEN err_msg = err_msg & vbCr &  "Please enter an employment services provider."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."		
LOOP UNTIL err_msg = ""

		

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

'Updating the ES Database
IF collecting_ES_statistics = true THEN
	for each member in member_array
		ES_Active = "Yes" 'Assigning new referral values for database
		ES_Counselor = "Unassigned"
		ESPrimary_Activity = "No Employment Plan"
		hh_member = member
		'msgbox case_number & hh_member & ESMemb_Name & EsSanction_Percentage & ESEmps_Status & ESTANF_MosUsed & ESExtension_Reason & ESDisa_End & ESPrimary_Activity & ESDate & ESprovider & ES_Counselor & ES_active & insert_string
		call write_MAXIS_info_to_ES_database(case_number, hh_member, ESMemb_Name, EsSanction_Percentage, ESEmps_Status, ESTANF_MosUsed, ESExtension_Reason, ESDisa_End, ESPrimary_Activity, ESDate, ESprovider, ES_Counselor, ES_active, insert_string)
		ESProvider = ES_provider 'resetting variables for other members
		ESDate = referral_date
	next
END IF

script_end_procedure("")

	
