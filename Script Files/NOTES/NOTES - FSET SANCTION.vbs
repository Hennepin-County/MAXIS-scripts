'OPTION EXPLICIT

name_of_script = "NOTES - FSET SANCTION.vbs"
start_time = timer

'DIM name_of_script
'DIM start_time
'DIM FuncLib_URL
'DIM run_locally
'DIM default_directory
'DIM beta_agency
'DIM req
'DIM fso
'DIM row
'DIM col

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
''END OF GLOBAL VARIABLES----------------------------------------------------------------------------------------------------


''SNAP_sanction_type_droplist dialog and other variables----------------------------------------------------------------------------------------------------
'DIM ButtonPressed
'DIM SNAP_sanction_type_dialog
'DIM case_number
'DIM footer_month
'DIM MAXIS_footer_month
'DIM footer_year
'DIM MAXIS_footer_year
'DIM worker_signature
'DIM sanction_type_droplist
'DIM ABAWD_status_check
'DIM WREG_MEMB_check
'MAXIS_footer_finder(MAXIS_footer_year, MAXIS_footer_month)
''SNAP_sanction_imposed_dialog
'DIM SNAP_sanction_imposed_dialog
'DIM sanction_begin_date
'DIM HH_Member_Number
'DIM PWE_check
'DIM number_of_sanction_droplist
'DIM sanction_reason_droplist
'DIM other_sanction_notes
'DIM agency_informed_sanction
'DIM WREG_sanction_droplist
''SNAP_sanction_resolved_dialog
'DIM SNAP_sanction_resolved_dialog
'DIM sanction_end_date
'DIM resolved_HH_Member_Number
'DIM resolved_PWE_check
'DIM sanction_resolved_reason_droplist
'DIM sanction_resolution_droplist
'DIM other_resolved_sanction_notes
'DIM ABAWD_status_droplist
'DIM Exempt_FSET_WREG_droplist
'DIM mandatory_WREG_exempt_FSET_droplist
'DIM FSET_orientation_date
'DIM orientation_letter_check
'DIM GA_basis_droplist


'FUNCTION----------------------------------------------------------------------------------------------------
FUNCTION MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)'Grabbing the footer month/year
	Call find_variable("Month: ", MAXIS_footer_month, 2)
	If isnumeric(MAXIS_footer_month) = true then 	'checking to see if a footer month 'number' is present 
		footer_month = MAXIS_footer_month	
		call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
		If isnumeric(MAXIS_footer_year) = true then footer_year = MAXIS_footer_year 'checking to see if a footer year 'number' is present
	Else 'If we don’t have one found, we’re going to assign the current month/year.
		MAXIS_footer_month = DatePart("m", date)   'Datepart delivers the month number to the variable
		If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month   'If it’s a single digit month, add a zero
		MAXIS_footer_year = right(DatePart("yyyy", date), 2)	'We only need the right two characters of the year for MAXIS
	End if
END FUNCTION

'The DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog SNAP_sanction_type_dialog, 0, 0, 171, 110, "SNAP Sanction type dialog					"
  EditBox 65, 10, 65, 15, case_number
  EditBox 65, 30, 30, 15, MAXIS_footer_month
  EditBox 100, 30, 30, 15, MAXIS_footer_year
  DropListBox 20, 65, 120, 15, "Select one..."+chr(9)+"Imposing sanction"+chr(9)+"Resolving sanction", sanction_type_droplist
  ButtonGroup ButtonPressed
    OkButton 25, 85, 50, 15
    CancelButton 80, 85, 50, 15
  Text 10, 30, 50, 15, "Footer month:"
  Text 10, 10, 50, 15, "Case number: "
  Text 5, 50, 175, 10, "Are you imposing or resolving the FSET sanction?"
EndDialog


BeginDialog SNAP_sanction_imposed_dialog, 0, 0, 351, 170, "SNAP sanction imposed dialog"
  EditBox 95, 5, 55, 15, sanction_begin_date
  EditBox 210, 5, 20, 15, HH_Member_Number
  CheckBox 240, 10, 110, 10, "Sanctioned individual is PWE", PWE_check
  DropListBox 90, 25, 255, 15, "Select one..."+chr(9)+"1st (1 month or until compliance, whichever is longer)"+chr(9)+"2nd (3 months or until compliance, whichever is longer)"+chr(9)+"3rd (6 months or until compliance, whichever is longer)", number_of_sanction_droplist
  DropListBox 90, 45, 255, 15, "Select one..."+chr(9)+"Failed to comply with SNAP E&T requirements"+chr(9)+"Failed to accept suitable employment w/o good cause"+chr(9)+"Voluntarily quit suitable employment w/o good cause"+chr(9)+"Voluntarily reduced work hours w/o good cause", sanction_reason_droplist
  EditBox 90, 65, 255, 15, other_sanction_notes
  EditBox 130, 85, 50, 15, agency_informed_sanction
  EditBox 150, 105, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 105, 50, 15
    CancelButton 295, 105, 50, 15
  Text 5, 10, 85, 10, "FSET sanction begin date:"
  Text 5, 70, 70, 10, "Other sanction notes:"
  Text 155, 10, 50, 10, "HH Member #:"
  Text 5, 90, 125, 10, "Date agency was notified of sanction:"
  Text 85, 110, 60, 10, "Worker signature:"
  Text 5, 135, 340, 25, "If client is PWE the ENTIRE unit is sanctioned.  If they are not the PWE, ONLY the member is sanctioned.  Also ABAWDs have until the end of the month prior to the effective date of the SNAP closing to cooperate with the SNAP E and T orientation/work requirements.  "
  Text 5, 30, 70, 10, "Number of sanctions:"
  Text 5, 50, 80, 10, "Reason for the sanction:"
  GroupBox 0, 125, 350, 40, "Per CM.0028.30.03"
EndDialog


BeginDialog SNAP_sanction_resolved_dialog, 0, 0, 346, 250, "SNAP sanction resolved dialog"
  EditBox 90, 10, 55, 15, sanction_end_date
  EditBox 200, 10, 20, 15, resolved_HH_Member_Number
  CheckBox 230, 15, 110, 10, "Sanctioned individual is PWE", resolved_PWE_check
  DropListBox 90, 30, 245, 15, "Select one..."+chr(9)+"Member served minimum sanction & verbally agrees to comply with SNAP E&T "+chr(9)+"Member leaves the unit's home"+chr(9)+"Member becomes exempt (work registration or E & T)", sanction_resolved_reason_droplist
  EditBox 90, 50, 245, 15, other_resolved_sanction_notes
  DropListBox 90, 70, 245, 15, "Select one..."+chr(9)+"01 Work Reg Exmpt"+chr(9)+"02 Under Age 18  "+chr(9)+"03 Age 50 Or Over"+chr(9)+"04 Caregiver Of  Minor Child *  "+chr(9)+"05 Pregnant      "+chr(9)+"06 Employed Avg Of 20 Hrs/Wk  "+chr(9)+"07 Wrk Experience Participant   "+chr(9)+"08 Other E & T Services *"+chr(9)+"09 Resides In A Waivered Area "+chr(9)+"10 ABAWD Counted Month         "+chr(9)+"11 2nd-3rd Month Period Of Elig"+chr(9)+"12 RCA Or GA Recipient     "+chr(9)+"13 ABAWD Extension", ABAWD_status_droplist
  DropListBox 145, 130, 190, 15, "Select one..."+chr(9)+"03 Temp/Perm Incap (Min 30 Days)"+chr(9)+"04  Responsible For Care Incap HH MEMB *"+chr(9)+"05 Age 60 or older"+chr(9)+"06 Under age 16"+chr(9)+"07 Age 16-17 living w/ parent/caregiver"+chr(9)+"08 Resp for care of child <6 *"+chr(9)+"09 Empl 30 hr/wk or earining = to min wage x 30 hr/wk"+chr(9)+"10 Matching grant participant"+chr(9)+"11 Receiving or applied for unemployment"+chr(9)+"12 Enrolled in school, training program or higher education"+chr(9)+"13 Participating In CD Program"+chr(9)+"14 Receiving MFIP"+chr(9)+"20 Pending/Receiving DWP Or WB"+chr(9)+"22 Applied for SSI", Exempt_FSET_WREG_droplist
  DropListBox 145, 150, 190, 15, "Select one..."+chr(9)+"15  Age 16-17 Not Lvg W/Pare/Crgvr"+chr(9)+"16  50-59 years old"+chr(9)+"21  Resp For Care Of Child < 18"+chr(9)+"17  Receiving RCA Or GA"+chr(9)+"18  Providing In-Home Schooling"+chr(9)+"30  Mandatory FSET participant", mandatory_WREG_exempt_FSET_droplist
  EditBox 145, 170, 55, 15, FSET_orientation_date
  CheckBox 210, 175, 125, 10, "SNAP E and T orientation letter sent", orientation_letter_check
  DropListBox 145, 190, 190, 15, "Select one..."+chr(9)+"04 Permanent Ill Or Incap"+chr(9)+"05 Temporary Ill Or Incap"+chr(9)+"06 Care Of Ill Or Incap Mbr"+chr(9)+"07 Resident Of Facility"+chr(9)+"08 Family Violence Indc"+chr(9)+"09 Mntl Ill Or Dev Disabled"+chr(9)+"10 SSI/RSDI Pend "+chr(9)+"11 Appealing SSI/RSDI Denial"+chr(9)+"12 Advanced Age"+chr(9)+"13 Learning Disability"+chr(9)+"15 Pregnant, 3rd Trimester"+chr(9)+"17 Protect/Court Ordered"+chr(9)+"20 Age 16 Or 17 SS Approval"+chr(9)+"25 Emancipated Minor"+chr(9)+"28 Unemployable"+chr(9)+"29 Displaced Hmkr (Ft Student)"+chr(9)+"30 Minor W/ Adult Unrelated"+chr(9)+"32 ESL, Adult/HS At Least Half Time, Adult"+chr(9)+"99 No Elig Basis", GA_basis_droplist
  EditBox 165, 230, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 230, 230, 50, 15
    CancelButton 285, 230, 50, 15
  Text 100, 235, 60, 10, "Worker signature:"
  Text 150, 15, 50, 10, "HH Member #:"
  Text 5, 135, 135, 10, "Exempt from FSET/WREG:"
  Text 5, 155, 125, 10, "Mandatory WREG, FSET (non)exempt:"
  Text 5, 35, 70, 10, "Sanction resolution: "
  GroupBox 0, 0, 340, 105, ""
  GroupBox 0, 115, 340, 95, "New FSET/WREG status, complete the one of the next two sections that applies to the member's status"
  Text 5, 55, 70, 10, "Other sanction notes:"
  Text 5, 175, 135, 10, "New FSET orientation date (if applicable):"
  Text 5, 75, 55, 10, "ABAWD Status:"
  Text 90, 90, 190, 10, "* in ABAWD status = 1 ABAWD exemption per household "
  Text 5, 195, 135, 10, "If GA basis exists then select  the basis:"
  Text 5, 15, 80, 10, "FSET sanction end date: "
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""
'Grabbing the case number
Call MAXIS_case_number_finder(case_number)
'Grabbing the footer month and footer year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'Initial dialog giving the user the option to select the type of sanction (imposed or resolved)
Do	
	Do	
		Do
			dialog SNAP_sanction_type_dialog
			cancel_confirmation
			If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
		Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8
		IF MAXIS_footer_month = "" OR MAXIS_footer_year = "" THEN MsgBox "You must enter both the footer month & footer year."
	LOOP until (MAXIS_footer_month <> "" AND MAXIS_footer_year <> "")
	IF sanction_type_droplist = "Select one..." THEN MsgBox "You must select either ""imposing sanction"" or ""resolving sanction""."
LOOP until sanction_type_droplist <> "Select one..."
'If worker selects to impose a sanction, they will get this dialog 
If sanction_type_droplist = "Imposing sanction" THEN
DO					
	DO				
		DO 			
			DO
				DO
					Do
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
ELSEIf sanction_type_droplist = "Resolving sanction" THEN	
	DO
		Do
			DO
				DO	
					DO
						DO
							dialog SNAP_sanction_resolved_dialog
							cancel_confirmation
							If sanction_end_date = "" THEN MsgBox "You must enter the date the sanction ends."
						LOOP until sanction_end_date <> ""
						If resolved_HH_Member_Number = "" THEN MsgBox "You must enter the client's member number"
					LOOP until resolved_HH_Member_Number <> ""
					If sanction_resolved_reason_droplist = "Select one..." THEN MsgBox "You must choose the reason the sanction has been resolved."
				LOOP until sanction_resolved_reason_droplist <> "Select one..."
				If ABAWD_status_droplist = "Select one..." THEN MsgBox "You must enter the ABAWD status."
			LOOP until ABAWD_status_droplist <> "Select one..."
			If worker_signature = "" THEN MsgBox "You must sign your case note."
		LOOP until worker_signature <> ""
		If(Exempt_FSET_WREG_droplist <> "Select one..." AND mandatory_WREG_exempt_FSET_droplist <> "Select one...") OR (Exempt_FSET_WREG_droplist = "Select one..." AND mandatory_WREG_exempt_FSET_droplist = "Select one...") THEN MsgBox "You must select only one of the options for the new FSET/WREG status."
	LOOP until (Exempt_FSET_WREG_droplist = "Select one..." AND mandatory_WREG_exempt_FSET_droplist <> "Select one...") OR (Exempt_FSET_WREG_droplist <> "Select one..." AND mandatory_WREG_exempt_FSET_droplist = "Select one...")
END If


'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Logic to have full policy verbiage in the case note (droplist could not support full policy verbiage)
IF sanction_resolved_reason_droplist = "Member served minimum sanction & verbally agrees to comply" THEN sanction_resolution_droplist = "Member served the minimum sanction period, and verbally agrees to comply with SNAP E&T during the SNAP application process." 

Call start_a_blank_CASE_NOTE
'Next 2 lines create custom headers based on the type of sanction chosen 
'Case note if imposing sanction
If sanction_type_droplist = "Imposing sanction" THEN 
	Call write_variable_in_CASE_NOTE("--SNAP sanction imposed for MEMB " & HH_Member_Number & ", eff: " & sanction_begin_date & "--")
	Call write_bullet_and_variable_in_CASE_NOTE("HH MEMB #", HH_Member_Number)
	If PWE_check = 1 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is the PWE. Entire household is sanctioned.")
	If PWE_check = 0 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is NOT the PWE. Only the HH MEMB is sanctioned.")
	Call write_bullet_and_variable_in_CASE_NOTE("Date agency was notified of sanction", agency_informed_sanction)
	Call write_bullet_and_variable_in_CASE_NOTE("Number/occurrence of sanction", number_of_sanction_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for sanction", sanction_reason_droplist)
	IF other_sanction_notes <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("Other sanction notes", other_sanction_notes)
	Call write_variable_in_CASE_NOTE("* Sanction WREG status: 02 Fail To Cooperate With FSET")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
'Case note if resolving sanction
ELSEIF sanction_type_droplist = "Resolving sanction" THEN
	Call write_variable_in_CASE_NOTE("--SNAP sanction ended for MEMB " & resolved_HH_Member_Number & ", eff: " & sanction_end_date & "--")
	Call write_bullet_and_variable_in_CASE_NOTE("HH MEMB #", resolved_HH_Member_Number)
	If resolved_PWE_check = 1 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is the PWE. Entire household's sanction is resolved.")
	IF resolved_PWE_check = 0 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is NOT the PWE. Only this HH MEMB's sanction is resolved.")
	Call write_bullet_and_variable_in_CASE_NOTE("Sanction resolution reason", sanction_resolved_reason_droplist)
	If other_resolved_sanction_notes <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("Other sanction notes", other_resolved_sanction_notes)
	Call write_variable_in_CASE_NOTE("=======WORK REGISTRATION INFO=======")
	IF Exempt_FSET_WREG_droplist <> "Select one..." THEN Call write_bullet_and_variable_in_CASE_NOTE("New FSET Work Reg Status", Exempt_FSET_WREG_droplist)
	IF mandatory_WREG_exempt_FSET_droplist <> "Select one..." THEN Call write_bullet_and_variable_in_CASE_NOTE("New FSET Work Reg Status", mandatory_WREG_exempt_FSET_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("New ABAWD status", ABAWD_status_droplist)
	IF GA_basis_droplist <> "Select one..." THEN Call write_bullet_and_variable_in_CASE_NOTE("New GA basis", GA_basis_droplist)
	IF FSET_orientation_date <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("New FSET orientation date", FSET_orientation_date)
	IF orientation_letter_check = 1 THEN Call write_variable_in_CASE_NOTE("* SNAP E&T orientation letter was sent to the client.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)	
END IF
PF3
PF3


'LOGIC----------------------------------------------------------------------------------------------------
'Logic to change the number_of_sanction_droplist into correct coding for the WREG panel
IF number_of_sanction_droplist = "1st (1 month or until compliance, whichever is longer)" then number_of_sanction_droplist = "01"
IF number_of_sanction_droplist = "2nd (3 months or until compliance, whichever is longer)" then number_of_sanction_droplist = "02"
IF number_of_sanction_droplist = "3rd (6 months or until compliance, whichever is longer)" then number_of_sanction_droplist = "03"

'Logic to change the GA_basis_droplist into correct coding for the WREG panel
GA_basis_droplist = Left(GA_basis_droplist, 2)

CAll MAXIS_background_check
'UPDATING THE WREG PANEL----------------------------------------------------------------------------------------------------
'Updates WREG if sanction is imposed
If sanction_type_droplist = "Imposing sanction" THEN 
	Call navigate_to_MAXIS_screen("STAT", "WREG")
	EMWriteScreen HH_Member_Number, 20, 76
	transmit
	'checking to make sure that WREG is updating for the correct member
	EMReadScreen WREG_MEMB_check, 7, 12, 2
	IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER " THEN script_end_procedure ("The member number that you entered is not valid.  Please check the member number, and start the script again.")
	'if MEMB number is correct the WREG is updated
	PF9
	EMWriteScreen "02", 8, 50
	Call create_MAXIS_friendly_date(sanction_begin_date, 0, 10, 50)
	EMWriteScreen number_of_sanction_droplist, 11, 50
	EMWriteScreen "_", 8, 80
	'Updates WREG if sanction is resolved	
ELSEif sanction_type_droplist = "Resolving sanction" THEN
	Call navigate_to_MAXIS_screen("STAT", "WREG")
	'checking to make sure HH MEMB is valid
	EMWriteScreen resolved_HH_Member_Number, 20, 76
	transmit
	EMReadScreen WREG_MEMB_check, 7, 12, 2
	IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER " THEN script_end_procedure ("The member number that you entered is not valid.  Please check the member number, and start the script again.") 
	'if MEMB number is correct the WREG is updated
	PF9
	IF Exempt_FSET_WREG_droplist <> "Select one..." THEN EMWriteScreen Exempt_FSET_WREG_droplist, 8, 50
	IF mandatory_WREG_exempt_FSET_droplist <> "Select one..." THEN EMWriteScreen mandatory_WREG_exempt_FSET_droplist, 8, 50
	If FSET_orientation_date <> "" THEN Call create_MAXIS_friendly_date(FSET_orientation_date, 0, 9, 50)
	EMWriteScreen "______", 10, 50 'deletes out the sanction date
	EMWriteScreen "______", 10, 53
	EMWriteScreen "______", 10, 56
	EMWriteScreen ABAWD_status_droplist, 13, 50 ' updates the ABAWD status 
	'updating the Defer FSET/No Funds (Y/N) field on WREG
	EMReadScreen ABAWD_status_check, 2, 13, 50	'checking for the coding on ABAWD status field
	IF ABAWD_status_check = "10" THEN 
		EMWriteScreen "N", 8, 80
	ELSE 
		EMWriteScreen "_", 8, 80
	END IF  
	'updates GA basis if GA basis exists
	If GA_basis_droplist <> "Se" THEN 
		EMWritescreen GA_basis_droplist, 15, 50	
	ELSE 
		EMWriteScreen "__", 15, 50
	END IF	
END IF

'end message to worker reminding users to check notices
script_end_procedure("Success, your case note been made and the WREG panel updated. Remember to approve your new results, and check your notice for accuracy.")