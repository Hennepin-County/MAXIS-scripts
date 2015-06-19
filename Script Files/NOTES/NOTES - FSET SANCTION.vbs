OPTION EXPLICIT

name_of_script = "NOTES - FSET SANCTION.vbs"
start_time = timer

DIM name_of_script
DIM start_time
DIM FuncLib_URL
DIM run_locally
DIM default_directory
DIM beta_agency
DIM req
DIM fso
DIM row

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
'END OF GLOBAL VARIABLES----------------------------------------------------------------------------------------------------

DIM ButtonPressed
DIM SNAP_sanction_type_dialog
DIM case_number
DIM footer_month
DIM MAXIS_footer_month
DIM footer_year
DIM MAXIS_footer_year
DIM worker_signature
DIM sanction_type_droplist
DIM SNAP_sanction_imposed_dialog
DIM HH_Member_Number
DIM PWE_check
DIM number_of_sanction_droplist
DIM sanction_reason_droplist
DIM other_sanction_notes
DIM agency_informed_sanction
DIM sanction_begin_date
DIM imposed_update_WREG_check
DIM SNAP_sanction_resolved_dialog
DIM resolved_HH_Member_Number
DIM resolved_PWE_check
DIM sanction_resolution_droplist
DIM resolved_other_sanction_notes
DIM sanction_end_date
DIM resolved_update_WREG_check


BeginDialog SNAP_sanction_type_dialog, 0, 0, 171, 110, "SNAP Sanction type dialog					"
  EditBox 65, 10, 65, 15, case_number
  EditBox 65, 30, 30, 15, MAXIS_footer_month
  EditBox 100, 30, 30, 15, MAXIS_footer_year
  DropListBox 20, 65, 120, 15, "Select one..."+chr(9)+"Imposing sanction "+chr(9)+"Resolving sanction", sanction_type_droplist
  ButtonGroup ButtonPressed
    OkButton 25, 85, 50, 15
    CancelButton 80, 85, 50, 15
  Text 10, 30, 50, 15, "Footer month:"
  Text 10, 10, 50, 15, "Case number: "
  Text 5, 50, 175, 10, "Are you imposing or resolving the FSET sanction?"
EndDialog


BeginDialog SNAP_sanction_imposed_dialog, 0, 0, 346, 160, "SNAP sanction imposed dialog"
  EditBox 95, 5, 55, 15, sanction_begin_date
  EditBox 210, 5, 20, 15, HH_Member_Number
  CheckBox 240, 10, 110, 10, "Sanctioned individual is PWE", PWE_check
  DropListBox 90, 25, 255, 15, "Select one..."+chr(9)+"1st  (1 month or until compiance, whichever is longer)"+chr(9)+"2nd (3 months or until compiance, whichever is longer)"+chr(9)+"3rd  (6 months or until compiance, whichever is longer)", number_of_sanction_droplist
  DropListBox 90, 45, 255, 15, "Select one..."+chr(9)+"Failed to attend SNAP overview"+chr(9)+"Failed to accept suitable employment w/o good cause"+chr(9)+"Voluntarily quit suitable employment w/o good cause"+chr(9)+"Voluntarily reduced work hours w/o good cause", sanction_reason_droplist
  EditBox 90, 65, 255, 15, other_sanction_notes
  EditBox 130, 85, 50, 15, agency_informed_sanction
  CheckBox 190, 90, 150, 10, "WREG has been updated to reflect sanction", imposed_update_WREG_check
  EditBox 145, 105, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 235, 105, 50, 15
    CancelButton 290, 105, 50, 15
  Text 5, 50, 80, 10, "Reason for the sanction:"
  Text 5, 10, 85, 10, "FSET sanction begin date:"
  Text 155, 10, 50, 10, "HH Member #:"
  Text 5, 70, 70, 10, "Other sanction notes:"
  Text 5, 90, 125, 10, "Date agency was notified of sanction:"
  Text 80, 110, 60, 10, "Worker signature:"
  Text 5, 130, 340, 25, "**Per CM 0028.30.06:  If client is PWE the ENTIRE unit is sanctioned.  If they are not the PWE, ONLY the member is sanctioned.  Also ABAWDs have until the end of the month prior to the effective date of the SNAP closing to cooperate with the SNAP E and T orientation/work requirements.  "
  Text 5, 30, 70, 10, "Number of sanctions:"
EndDialog


BeginDialog SNAP_sanction_resolved_dialog, 0, 0, 336, 100, "SNAP sanction resolved dialog"
  EditBox 90, 5, 55, 15, sanction_end_date
  EditBox 200, 5, 20, 15, resolved_HH_Member_Number
  CheckBox 230, 10, 110, 10, "Sanctioned individual is PWE", resolved_PWE_check
  DropListBox 80, 25, 250, 15, "Select one..."+chr(9)+"Member served minimum sanction & verbally agrees to comply"+chr(9)+"Member leaves the unit's home"+chr(9)+"Member becomes exempt (work registration or E & T)", sanction_resolution_droplist
  EditBox 80, 45, 250, 15, resolved_other_sanction_notes
  CheckBox 50, 65, 280, 10, "WREG panel has been updated to reflect new status (no longer 'in sanction' status)", resolved_update_WREG_check
  EditBox 135, 80, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 80, 50, 15
    CancelButton 280, 80, 50, 15
  Text 150, 10, 50, 10, "HH Member #:"
  Text 5, 10, 85, 10, "FSET sanction end date: "
  Text 5, 50, 70, 10, "Other sanction notes:"
  Text 5, 30, 70, 10, "Sanction resolution: "
  Text 75, 85, 60, 10, "Worker signature:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""
'Grabbing the case number
Call MAXIS_case_number_finder(case_number)

'Grabbing the footer month/year
Call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
	footer_month = MAXIS_footer_month
	call find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
	If row <> 0 then footer_year = MAXIS_footer_year
End if

'Initial dialog giving the user the option to select the type of sanction (imposed or resolved)

Do	
	Do
		Do
			Dialog SNAP_sanction_type_dialog
			cancel_confirmation
			If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
		Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
		IF MAXIS_footer_month = "" OR MAXIS_footer_year = "" THEN MsgBox "You must enter both the footer month & footer year."
	LOOP until (MAXIS_footer_month <> "" AND MAXIS_footer_year <> "")
	If sanction_type_droplist = "Select one..." THEN MsgBox "You must choose to either impose or resolve the sanction."
LOOP until sanction_type_droplist <> "Select one..."

'If worker selects to impose a sanction, they will get this dialog 
If sanction_type_droplist = "Imposing sanction" THEN
	DO
		DO					
			DO				
				DO 			
					DO
						DO
							dialog SNAP_sanction_imposed_dialog
							cancel_confirmation
							If sanction_begin_date = "" THEN MsgBox "You must enter the date the sanction begins."
						LOOP until sanction_begin_date <> ""
						If HH_Member_Number = "" THEN MsgBox "You must enter the client's member number"
					LOOP until HH_Member_Number <> ""
					If number_of_sanction_droplist = "Select one..." THEN MsgBox "You must choose the number of sanctions."
				LOOP until number_of_sanction_droplist <> "Select one..."
				If sanction_reason_droplist = "Select one..." THEN MsgBox "You must choose the reason for the sanction."
			LOOP until sanction_reason_droplist <> "Select one..."
			If agency_informed_sanction = "" THEN MsgBox "You must enter the date the agency was informed of the sanction."
		LOOP until agency_informed_sanction <> ""
		If worker_signature = "" THEN MsgBox "You must sign your case note."
	LOOP until worker_signature <> ""
	'If worker selects to resolve a sanction, they will get this dialog
	ELSE If sanction_type_droplist = "Resolving sanction" THEN
	
		DO
			DO
				DO	
					DO
						dialog SNAP_sanction_resolved_dialog
						cancel_confirmation
						If sanction_end_date = "" THEN MsgBox "You must enter the date the sanction ends."
					LOOP until sanction_end_date <> ""
					If HH_Member_Number = "" THEN MsgBox "You must enter the client's member number"
				LOOP until HH_Member_Number <> ""
				If sanction_resolution_droplist = "Select one..." THEN MsgBox "You must choose the reason the sanction has been resolved."
			LOOP until sanction_resolution_droplist <> "Select one..."
			If worker_signature = "" THEN MsgBox "You must sign your case note."
		LOOP until worker_signature <> ""
	END IF 	
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Next 2 lines create custom headers based on the type of sanction chosen 
'Case note if imposing sanction
If sanction_type_droplist = "Imposing sanction" THEN 
	Call write_variable_in_CASE_NOTE("--Imposing SNAP sanction for MEMB " & HH_Member_Number & ", effective:" & sanction_begin_date & "--")
	Call write_bullet_and_variable_in_CASE_NOTE("FSET sanction begin date:", sanction_begin_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Household member number", HH_Member_Number)
		If PWE_check = 1 THEN Call write_bullet_and_variable_in_CASE_NOTE("Sanctioned individual is the PWE. Entire household is sanctioned.")
			ELSE Call write_bullet_and_variable_in_CASE_NOTE("Sanctioned individual is NOT the PWE. Only", HH_Member_Number & "is sanctioned.")
		END IF
	Call write_bullet_and_variable_in_CASE_NOTE("Number/occurrence of sanction", number_of_sanction_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for sanction", sanction_reason_droplist)
	IF other_sanction_notes <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("Other sanction notes", other_sanction_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("Date agency was notified of sanction", agency_informed_sanction)
	If imposed_update_WREG_check = 1 THEN Call write_bullet_and_variable_in_CASE_NOTE ("The WREG panel has been updated to reflect sanction.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
'Case note if resolving sanction	
'	ELSE 
'		IF sanction_type_droplist = "Resolving sanction" THEN 
'		Call write_variable_in_CASE_NOTE("--Resolving SNAP sanction for MEMB " & HH_Member_Number & ", effective:" & sanction_end_date & "--")
'		Call write_bullet_and_variable_in_CASE_NOTE("FSET sanction end date", sanction_end_date)
'		Call write_bullet_and_variable_in_CASE_NOTE("Household member number", resolved_HH_Member_Number)
'			If resolved_PWE_check = 1 THEN Call write_bullet_and_variable_in_CASE_NOTE("Sanctioned individual is the PWE. Entire household's sanction is resolved.")
'				ELSE IF Call write_bullet_and_variable_in_CASE_NOTE("Sanctioned individual is NOT the PWE. Only" & HH_Member_Number & "'s sanction is resolved.")
'			End If
'		Call write_bullet_and_variable_in_CASE_NOTE("Sanction resolution reason", sanction_resolution_droplist)
'		If resolved_ other_sanction_notes <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("Other sanction notes", resolved_ other_sanction_notes)
'		If resolved_update_WREG_check = 1 THEN Call write_bullet_and_variable_in_CASE_NOTE("The WREG panel has been updated to reflect new status (no longer 'in sanction' status).")
'		Call write_variable_in_CASE_NOTE("---")
'		Call write_variable_in_CASE_NOTE(worker_signature)	

script_end_procedure("")

