'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - FSET SANCTION.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 193                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("12/30/2016", "Updated reason for sanction reasons for newly added 'reason for sanction' field on STAT/WREG. Also updated user dialogs to include sanction month/year instead of full sanction date.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog SNAP_sanction_type_dialog, 0, 0, 181, 75, "SNAP Sanction type dialog"
  EditBox 95, 10, 50, 15, MAXIS_case_number
  DropListBox 95, 30, 80, 15, "Select one..."+chr(9)+"Imposing sanction"+chr(9)+"Resolving sanction", sanction_type_droplist
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 95, 50, 50, 15
  Text 45, 15, 45, 10, "Case number:"
  Text 10, 35, 80, 10, "Select a sanction option:"
EndDialog

BeginDialog SNAP_sanction_imposed_dialog, 0, 0, 306, 205, "SNAP sanction imposed dialog"
  EditBox 125, 10, 20, 15, MAXIS_footer_month
  EditBox 150, 10, 20, 15, MAXIS_footer_year
  EditBox 240, 10, 20, 15, HH_Member_Number
  EditBox 125, 35, 50, 15, agency_informed_sanction
  CheckBox 190, 40, 110, 10, "Sanctioned individual is PWE", PWE_check
  DropListBox 90, 60, 180, 15, "Select one..."+chr(9)+"1st (1 mo. or until compliance, whichever is longer)"+chr(9)+"2nd (3 mo. or until compliance, whichever is longer)"+chr(9)+"3rd (6 mo. or until compliance, whichever is longer)", number_of_sanction_droplist
  DropListBox 90, 80, 180, 15, "Select one..."+chr(9)+"(01) Failed to attend orientation"+chr(9)+"(02) Failed to develop work plan"+chr(9)+"(03) Failed to follow work plan", sanction_reason_droplist
  EditBox 90, 100, 180, 15, other_sanction_notes
  EditBox 75, 125, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 125, 50, 15
    CancelButton 240, 125, 50, 15
  Text 15, 105, 70, 10, "Other sanction notes:"
  Text 190, 15, 50, 10, "HH Member #:"
  Text 5, 40, 115, 10, "Agency was notified of sanction on:"
  Text 10, 130, 60, 10, "Worker signature:"
  Text 10, 160, 280, 35, "If client is PWE the ENTIRE unit is sanctioned.  If they are not the PWE, ONLY the member is sanctioned.  Also ABAWDs have until the end of the month prior to the effective date of the SNAP closing to cooperate with the SNAP E and T orientation/work requirements.  "
  Text 15, 65, 70, 10, "Number of sanctions:"
  Text 5, 85, 80, 10, "Reason for the sanction:"
  GroupBox 5, 145, 285, 55, "Per CM.0028.30.03"
  Text 5, 15, 110, 10, "FSET sanction begin month/year:"
EndDialog

BeginDialog SNAP_sanction_resolved_dialog, 0, 0, 346, 250, "SNAP sanction resolved dialog"
  EditBox 90, 10, 20, 15, MAXIS_footer_month
   EditBox 115, 10, 20, 15, MAXIS_footer_year
  EditBox 200, 10, 20, 15, resolved_HH_Member_Number
  CheckBox 230, 15, 110, 10, "Sanctioned individual is PWE", resolved_PWE_check
  DropListBox 90, 30, 245, 15, "Select one..."+chr(9)+"Member served minimum sanction & verbally agrees to comply with SNAP E&T "+chr(9)+"Member leaves the unit's home"+chr(9)+"Member becomes exempt (work registration or E & T)", sanction_resolved_reason_droplist
  EditBox 90, 50, 245, 15, other_resolved_sanction_notes
  DropListBox 90, 70, 245, 15, "Select one..."+chr(9)+"01 Work Reg Exmpt"+chr(9)+"02 Under Age 18  "+chr(9)+"03 Age 50 Or Over"+chr(9)+"04 Caregiver Of  Minor Child *  "+chr(9)+"05 Pregnant      "+chr(9)+"06 Employed Avg Of 20 Hrs/Wk  "+chr(9)+"07 Wrk Experience Participant   "+chr(9)+"08 Other E & T Services *"+chr(9)+"09 Resides In A Waivered Area "+chr(9)+"10 ABAWD Counted Month         "+chr(9)+"11 2nd-3rd Month Period Of Elig"+chr(9)+"12 RCA Or GA Recipient     "+chr(9)+"13 ABAWD Banked Months", ABAWD_status_droplist
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
  Text 5, 35, 70, 10, "Sanction resolution:"
  GroupBox 0, 0, 340, 105, ""
  GroupBox 0, 115, 340, 95, "New FSET/WREG status, complete the one of the next two sections that applies to the member's status"
  Text 5, 55, 70, 10, "Other sanction notes:"
  Text 5, 175, 135, 10, "New FSET orientation date (if applicable):"
  Text 5, 75, 55, 10, "ABAWD Status:"
  Text 90, 90, 190, 10, "* in ABAWD status = 1 ABAWD exemption per household "
  Text 5, 195, 135, 10, "If GA basis exists then select  the basis:"
  Text 5, 15, 85, 10, "Sanction end month/year:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing case number and footer month/year
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Initial dialog giving the user the option to select the type of sanction (imposed or resolved)
Do
	Do
		err_msg = ""
		dialog SNAP_sanction_type_dialog
		IF buttonpressed = 0 then stopscript 
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF sanction_type_droplist = "Select one..." then err_msg = err_msg & vbnewline & "* Select a sanction option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'If worker selects to impose a sanction, they will get this dialog
If sanction_type_droplist = "Imposing sanction" THEN
	DO
		Do
			err_msg = ""
			dialog SNAP_sanction_imposed_dialog
			cancel_confirmation
			If isnumeric(MAXIS_footer_month) = false then err_msg = err_msg & vbnewline & "* You must enter the footer month to begin the sanction."
			If isnumeric(MAXIS_footer_year) = false then err_msg = err_msg & vbnewline & "* You must enter the footer year to begin the sanction."
			If isnumeric(HH_Member_Number) = False or len(HH_Member_Number) <> 2 then err_msg = err_msg & vbnewline & "* You must enter the client's two-digit member number"
			If number_of_sanction_droplist = "Select one..." then err_msg = err_msg & vbnewline & "* You must choose the number of sanctions."
			If sanction_reason_droplist = "Select one..." then err_msg = err_msg & vbnewline & "* You must choose the reason for the sanction."
			If isDate(agency_informed_sanction) = False then err_msg = err_msg & vbnewline & "* You must enter the date the agency was informed of the sanction."
			If worker_signature = "" then err_msg = err_msg & vbnewline & "* You must sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

'If worker selects to resolve a sanction, they will get this dialog
ELSEIf sanction_type_droplist = "Resolving sanction" THEN
	DO
		DO
			err_msg = ""
			dialog SNAP_sanction_resolved_dialog
			cancel_confirmation
			If isnumeric(MAXIS_footer_month) = false then err_msg = err_msg & vbnewline & "* You must enter the footer month to end the sanction."
			If isnumeric(MAXIS_footer_year) = false then err_msg = err_msg & vbnewline & "* You must enter the footer year to end the sanction."
			If isnumeric(resolved_HH_Member_Number) = false or len(resolved_HH_Member_Number) <> 2 then err_msg = err_msg & vbnewline & "* You must enter the client's two-digit member number."
			If sanction_resolved_reason_droplist = "Select one..." then err_msg = err_msg & vbnewline & "* You must choose the reason the sanction has been resolved."
			If ABAWD_status_droplist = "Select one..." then err_msg = err_msg & vbnewline & "* You must enter the ABAWD status."
			If worker_signature = "" then err_msg = err_msg & vbnewline & "* You must sign your case note."
			If(Exempt_FSET_WREG_droplist <> "Select one..." AND mandatory_WREG_exempt_FSET_droplist <> "Select one...") OR (Exempt_FSET_WREG_droplist = "Select one..." AND mandatory_WREG_exempt_FSET_droplist = "Select one...") then err_msg = err_msg & vbnewline & "* You must select only one of the options for the new FSET/WREG status."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Logic to have full policy verbiage in the case note (droplist could not support full policy verbiage)
IF sanction_resolved_reason_droplist = "Member served minimum sanction & verbally agrees to comply" THEN sanction_resolution_droplist = "Member served the minimum sanction period, and verbally agrees to comply with SNAP E&T during the SNAP application process."

start_a_blank_CASE_NOTE
'Case note if imposing sanction
If sanction_type_droplist = "Imposing sanction" THEN
	Call write_variable_in_CASE_NOTE("--SNAP sanction imposed for MEMB " & HH_Member_Number & " eff: " & MAXIS_footer_month & "/" & MAXIS_footer_year & "--")
	Call write_bullet_and_variable_in_CASE_NOTE("HH MEMB #", HH_Member_Number)
	If PWE_check = 1 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is the PWE. Entire household is sanctioned.")
	If PWE_check = 0 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is NOT the PWE. Only the HH MEMB is sanctioned.")
	Call write_bullet_and_variable_in_CASE_NOTE("Date agency was notified of sanction", agency_informed_sanction)
	Call write_bullet_and_variable_in_CASE_NOTE("Number/occurrence of sanction", number_of_sanction_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for sanction", sanction_reason_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("Other sanction notes", other_sanction_notes)
	Call write_variable_in_CASE_NOTE("* Sanction WREG status: 02 Fail To Cooperate With FSET")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
'Case note if resolving sanction
ELSEIF sanction_type_droplist = "Resolving sanction" THEN
	Call write_variable_in_CASE_NOTE("--SNAP sanction ended for MEMB " & resolved_HH_Member_Number & " eff: " & MAXIS_footer_month & "/" & MAXIS_footer_year & "--")
	Call write_bullet_and_variable_in_CASE_NOTE("HH MEMB #", resolved_HH_Member_Number)
	If resolved_PWE_check = 1 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is the PWE. Entire household's sanction is resolved.")
	IF resolved_PWE_check = 0 THEN Call write_variable_in_CASE_NOTE("* Sanctioned individual is NOT the PWE. Only this HH MEMB's sanction is resolved.")
	Call write_bullet_and_variable_in_CASE_NOTE("Sanction resolution reason", sanction_resolved_reason_droplist)
    Call write_bullet_and_variable_in_CASE_NOTE("Other sanction notes", other_resolved_sanction_notes)
	Call write_variable_in_CASE_NOTE("=======WORK REGISTRATION INFO=======")
	IF Exempt_FSET_WREG_droplist <> "Select one..." THEN Call write_bullet_and_variable_in_CASE_NOTE("New FSET Work Reg Status", Exempt_FSET_WREG_droplist)
	IF mandatory_WREG_exempt_FSET_droplist <> "Select one..." THEN Call write_bullet_and_variable_in_CASE_NOTE("New FSET Work Reg Status", mandatory_WREG_exempt_FSET_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("New ABAWD status", ABAWD_status_droplist)
	IF GA_basis_droplist <> "Select one..." THEN Call write_bullet_and_variable_in_CASE_NOTE("New GA basis", GA_basis_droplist)
	Call write_bullet_and_variable_in_CASE_NOTE("New FSET orientation date", FSET_orientation_date)
	IF orientation_letter_check = 1 THEN Call write_variable_in_CASE_NOTE("* SNAP E&T orientation letter was sent to the client.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
END IF
PF3
PF3

'Logic to change the number_of_sanction_droplist into correct coding for the WREG panel----------------------------------------------------------------------------------------------------
IF number_of_sanction_droplist = "1st (1 mo. or until compliance, whichever is longer)"  then number_of_sanction_droplist = "01"
IF number_of_sanction_droplist = "2nd (3 mo. or until compliance, whichever is longer)" then number_of_sanction_droplist = "02"
IF number_of_sanction_droplist = "3rd (6 mo. or until compliance, whichever is longer)" then number_of_sanction_droplist = "03"

IF sanction_reason_droplist = "(01) Failed to attend orientation" then sanction_reason_droplist = "01"
IF sanction_reason_droplist = "(02) Failed to develop work plan"  then sanction_reason_droplist = "02"
IF sanction_reason_droplist = "(03) Failed to follow work plan"   then sanction_reason_droplist = "03" 

'Logic to change the GA_basis_droplist into correct coding for the WREG panel
GA_basis_droplist = Left(GA_basis_droplist, 2)

'Updating WREG panel----------------------------------------------------------------------------------------------------
MAXIS_footer_month_confirmation                 'confirming that we are in the correct footer month/year

Call navigate_to_MAXIS_screen("STAT", "WREG")

'Updates WREG if sanction is imposed
If sanction_type_droplist = "Imposing sanction" THEN
	EMWriteScreen HH_Member_Number, 20, 76
	transmit
	'checking to make sure that WREG is updating for the correct member
	EMReadScreen WREG_MEMB_check, 6, 24, 2
	IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER" THEN script_end_procedure ("The member number that you entered is not valid.  Please check the member number, and start the script again.")
	'if MEMB number is correct the WREG is updated
	PF9
	EMWriteScreen "02", 8, 50							'Enters sanction FSET code of "02"
	EMWriteScreen MAXIS_footer_month, 10, 50			'sanction begin month
	EMWriteScreen MAXIS_footer_year, 10, 56			'sanction begin year
	EMWriteScreen number_of_sanction_droplist, 11, 50	'sanction #
	EMWriteScreen sanction_reason_droplist, 12, 50		'new field - reason for sanction. This adds information to the notice
	EMWriteScreen "_", 8, 80							'blanks out Defer FSET/No funds field 
	'Updates WREG if sanction is resolved
ELSEif sanction_type_droplist = "Resolving sanction" THEN
	'checking to make sure HH MEMB is valid
	EMWriteScreen resolved_HH_Member_Number, 20, 76
	transmit
	EMReadScreen WREG_MEMB_check, 6, 24, 2
	IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER" THEN script_end_procedure ("The member number that you entered is not valid.  Please check the member number, and start the script again.")
	'if MEMB number is correct the WREG is updated
	PF9
	IF Exempt_FSET_WREG_droplist <> "Select one..." THEN EMWriteScreen Exempt_FSET_WREG_droplist, 8, 50
	IF mandatory_WREG_exempt_FSET_droplist <> "Select one..." THEN EMWriteScreen mandatory_WREG_exempt_FSET_droplist, 8, 50
	If FSET_orientation_date <> "" THEN Call create_MAXIS_friendly_date(FSET_orientation_date, 0, 9, 50)
	EMWriteScreen "______", 10, 50 'deletes out the sanction date
	EMWriteScreen "______", 10, 56
	
	EMWriteScreen ABAWD_status_droplist, 13, 50 ' updates the ABAWD status
	'updating the Defer FSET/No Funds (Y/N) field on WREG
	EMReadScreen ABAWD_status_check, 2, 13, 50	'checking for the coding on ABAWD status field
	IF ABAWD_status_check = "10" or ABAWD_status_check = "13" THEN
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

script_end_procedure("Success, your case note been made and the WREG panel updated. Remember to approve your new results, and check your notice for accuracy.")