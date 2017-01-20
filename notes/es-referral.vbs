'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ES REFERRAL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
call changelog_update("01/17/2017", "Added program type (DWP or MFIP) into case note header.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

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

'DIALOGS----------------------------------------------------------------------------------------------------
  BeginDialog ES_referral_dialog, 0, 0, 286, 250, "Employment services referral"
   EditBox 70, 10, 50, 15, MAXIS_case_number
   DropListBox 230, 10, 50, 15, "Select one..."+chr(9)+"DWP"+chr(9)+"MFIP", select_program
   EditBox 70, 35, 50, 15, referral_date
   EditBox 230, 35, 50, 15, plan_deadline
   EditBox 230, 55, 50, 15, hh_member_list
   CheckBox 20, 75, 255, 10, "Set a TIKL for DWP employment plan date to deny DWP if plan not received.", TIKL_check
   DropListBox 70, 115, 60, 15, "Select one..."+chr(9)+"Scheduled"+chr(9)+"Rescheduled", appt_type
   EditBox 230, 115, 50, 15, appt_date
   IF collecting_ES_statistics = True THEN
    DropListBox  85, 90, 185, 15, ES_agency_list, ES_provider
   ELSE
    EditBox 70, 145, 210, 15, ES_provider
   END IF
   EditBox 70, 165, 210, 15, vendor_num
   EditBox 70, 185, 210, 15, other_notes
   CheckBox 20, 205, 135, 10, "DHS 4161 (DWP referral) sent to client", dwp_referral_check
   CheckBox 165, 205, 120, 10, "Paper referral sent to ES provider", es_referral_check
   EditBox 70, 225, 100, 15, worker_signature
   ButtonGroup ButtonPressed
     OkButton 175, 225, 50, 15
     CancelButton 230, 225, 50, 15
   Text 70, 60, 155, 10, "HH members referred (separate with commas):"
   Text 25, 150, 45, 10, "ES Provider:"
   Text 45, 190, 25, 10, "Notes:"
   Text 5, 230, 60, 10, "Worker Signature:"
   Text 15, 15, 50, 10, "Case Number:"
   Text 130, 35, 95, 20, "DWP employment plan deadline (10 Business days):"
   Text 30, 170, 40, 10, "Vendor #'s:"
   GroupBox 5, 95, 280, 45, "If the ES appt. has been scheduled/needs to be rescheduled:"
   Text 20, 40, 45, 10, "Referral Date:"
   Text 190, 15, 35, 10, "Program:"
   Text 10, 120, 60, 10, "Appointment type:"
   Text 135, 120, 90, 10, "Schedule/Reschedule date:"
 EndDialog

'Connecting to MAXIS and grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------Calling the dialog / requiring completion of most fields.
DO
    DO
        err_msg = ""
	    Dialog ES_referral_dialog
	    cancel_confirmation
	    IF isnumeric(MAXIS_case_number) = FALSE THEN err_msg = err_msg & vbCr &  "Please enter a valid case number."
        If select_program = "Select one..." THEN err_msg = err_msg & vbCr & "* Select the cash program."
        IF referral_date = "" THEN err_msg = err_msg & vbCr & "* Please enter a referral date."
	    if (appt_type <> "Select one..." and isdate(appt_date) = False) THEN err_msg = err_msg & vbCr & "* Please enter the appointment date."
        IF appt_date <> "" and appt_type = "Select one..." THEN err_msg = err_msg & vbCr & "* Please select the appointment type."
	    IF ES_provider = "" THEN err_msg = err_msg & vbCr &  "Please enter an employment services provider."
	    IF worker_signature = "" THEN err_msg = err_msg & vbCr &  "Please sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

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

'variable for case note header based on the appt type selected
If appt_type = "Select one..." then case_note_header = "**ES referral sent for " & select_program & "**"
If appt_type = "Scheduled" then case_note_header = "**" & select_program & " ES Appt Scheduled for " & appt_date & "**"
If appt_type = "Rescheduled" then case_note_header = "**" & select_program & " ES Appt Rescheduled for " & appt_date & "**"

'The case note --------- -------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE(case_note_header)
If appt_type = "Select one..." then 
	call write_variable_in_CASE_NOTE("* Referral sent to " & ES_provider & " on " & referral_date & ".")
Else 
	call write_bullet_and_variable_in_CASE_NOTE("ES provider", ES_provider)
End if 
call write_bullet_and_variable_in_CASE_NOTE("Members referred", hh_member_list)
IF select_program = "DWP" THEN call write_bullet_and_variable_in_CASE_NOTE("Employment plan due back on", plan_deadline)
IF dwp_referral_check = 1 THEN CALL write_variable_in_CASE_NOTE("* DHS 4161 sent to client.")
IF TIKL_check = 1 THEN CALL write_variable_in_CASE_NOTE("* TIKL to deny DWP this date if employment plan has not been completed.")
IF es_referral_check = 1 THEN CALL write_variable_in_CASE_NOTE("* Paper referral sent to ES provider.")
Call write_bullet_and_variable_in_CASE_NOTE("Vendor #(s)", vendor_num)
call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Updating the ES Database
IF collecting_ES_statistics = true THEN
	for each member in member_array
		ES_Active = "Yes" 'Assigning new referral values for database
		ES_Counselor = "Unassigned"
		ESPrimary_Activity = "No Employment Plan"
		hh_member = member
		'msgbox MAXIS_case_number & hh_member & ESMemb_Name & EsSanction_Percentage & ESEmps_Status & ESTANF_MosUsed & ESExtension_Reason & ESDisa_End & ESPrimary_Activity & ESDate & ESprovider & ES_Counselor & ES_active & insert_string
		call write_MAXIS_info_to_ES_database(MAXIS_case_number, hh_member, ESMemb_Name, EsSanction_Percentage, ESEmps_Status, ESTANF_MosUsed, ESExtension_Reason, ESDisa_End, ESPrimary_Activity, ESDate, ESprovider, ES_Counselor, ES_active, insert_string)
		ESProvider = ES_provider 'resetting variables for other members
		ESDate = referral_date
	next
END IF

script_end_procedure("")