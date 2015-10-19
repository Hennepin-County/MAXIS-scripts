'AUTHOR: ROBERT KALB AND CHARLES POTTER, ANOKA COUNTY
'Estimated manual processing time of 12 mins 30 seconds (if going through all exemptions with client)

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - ABAWD SCREENING TOOL.vbs"
start_time = timer

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

'Dialogs===================================================================================================================
'This dialog is for the WREG exemptions.-----------------------------------------------------------------------
BeginDialog wreg_exemptions, 0, 0, 311, 250, "ABAWD Screening Tool"
  CheckBox 5, 20, 305, 10, "...Permanently or Temporarily disabled or incapacitated (at least 30 days)? (CM0010.18.06)", wreg_disa
  CheckBox 5, 35, 270, 10, "...responsible for the care of a disabled household member?", care_of_hh_memb
  CheckBox 5, 50, 275, 10, "...age 60 or older?", age_sixty
  CheckBox 5, 65, 290, 15, "...under the age of 16?", under_sixteen
  CheckBox 5, 85, 275, 10, "...aged 16 or 17 living w/ parent or caregiver?", sixteen_seventeen
  CheckBox 5, 100, 275, 10, "...responsible for the care of a child under 6?", care_child_six
  CheckBox 5, 115, 255, 10, "...employed 30 hours per week or earning at least $935.25/month gross?", employed_thirty
  CheckBox 5, 130, 255, 10, "...receiving or applied for unemployment insurance?", unemployment
  CheckBox 5, 145, 255, 10, "...enrolled in school, training program, or higher education?", enrolled_school
  CheckBox 5, 160, 305, 10, "...participating in a chemical dependency treatment program (not AA or Narc. Anonymous)?", CD_program
  CheckBox 5, 175, 300, 10, "...receiving MFIP?", receiving_MFIP
  CheckBox 5, 190, 305, 10, "...receiving or pending for Diversionary Work Program or Work Benefit?", receiving_DWP_WB
  CheckBox 5, 205, 300, 10, "...applied for SSI (can be in first appeal)?", applied_SSI
  ButtonGroup ButtonPressed
    PushButton 205, 235, 50, 15, "NEXT", next_button
    CancelButton 260, 235, 50, 15
  Text 5, 5, 85, 10, "Is the client..."
EndDialog

'This dialog gets the client's case number.---------------------------------------------------------------------
BeginDialog get_case_number, 0, 0, 181, 100, "ABAWD Screening Tool"
  Text 10, 15, 50, 10, "Case Number: "
  EditBox 90, 10, 50, 15, case_number
  Text 10, 35, 70, 10, "Member Number:"
  EditBox 90, 30, 30, 15, member_number
  Text 10, 55, 75, 10, "Sign your Case Note:"
  EditBox 90, 50, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 45, 85, 50, 15, "Next", next_button
    CancelButton 95, 85, 50, 15
EndDialog


'This dialog is for the ABAWD exemptions and is used if the CL does not have a WREG exemption.---------------------
BeginDialog abawd_exemptions, 0, 0, 241, 180, "ABAWD Screening Tool"
  CheckBox 5, 20, 230, 15, "...residing in a waivered area", waiver
  CheckBox 5, 35, 185, 15, "...younger than 18 OR 50 or older?", age_exempt
  CheckBox 5, 50, 210, 15, "...medically certified as pregnant?", cert_preg
  CheckBox 5, 65, 210, 15, "...working at least 20 hours per week or 80 hours per month?", working_20
  CheckBox 5, 80, 230, 15, "...receiving RCA or GA?", receiving_cash
  CheckBox 5, 95, 240, 15, "...responsible for the care of a dependent child?", dependent_child
  CheckBox 5, 110, 240, 15, "...a Work Experience participant?", work_exp
  CheckBox 5, 125, 240, 15, "...participating in an approved Employment and Training program?", approved_ET
  ButtonGroup ButtonPressed
    PushButton 45, 160, 50, 15, "Previous", previous_button
    PushButton 100, 160, 50, 15, "Next", next_button
    CancelButton 180, 160, 50, 15
  Text 5, 5, 245, 10, "Is the client..."
EndDialog

'This dialog allows the screener to ask if the CL has earned an additional 3-month period of ABAWD-counted months---------
BeginDialog earn_additional_months, 0, 0, 366, 95, "ABAWD Screening Tool"
  CheckBox 5, 30, 355, 15, "Has the CL worked at least 80 hours in a month SINCE closing for using their last ABAWD-counted month?", worked_80_since_closing
  CheckBox 5, 50, 355, 15, "Has the CL used a second period of ABAWD-counted months?", has_used_second_period
  ButtonGroup ButtonPressed
    PushButton 165, 75, 50, 15, "Finish", finish_button
    PushButton 110, 75, 50, 15, "Previous", previous_button
    CancelButton 220, 75, 50, 15
  Text 5, 10, 295, 15, "Please navigate to the ABAWD Tracking Record for the appropriate member..."
EndDialog


'This dialog gets the worker's signature and allows the OSA to enter any comments for the case worker.----------------------
'The idea being that if the OSA notices irregularities or unusualness (word?) in the ABAWD tracking panel, it---------------
'can be reported to the worker or the worker can be directed to look deeper into the ABAWD tracking.------------------------
BeginDialog get_worker_comments, 0, 0, 166, 105, "ABAWD Screening Tool"
  EditBox 5, 50, 155, 15, worker_comment
  ButtonGroup ButtonPressed
    PushButton 20, 75, 50, 15, "OK", OK_button
    CancelButton 90, 75, 50, 15
  Text 5, 10, 150, 10, "Case noting CL interaction."
  Text 5, 25, 160, 20, "Any additional comments, please enter here. Press ENTER to complete and Case Note."
EndDialog

'FUNCTIONS========================================================================================
'Two functions were created
'One to count ABAWD months, it counts M and X months basing its search on a period of 3 years (36 months) since the WREG panel shifts as years go by.
function how_many_abawd_months(abawd_counted_months)
  call navigate_to_screen("stat", "wreg")
    EMWriteScreen member_number, 20, 76
    transmit
    EMSetCursor 13, 57
    EMSendKey "X"
    transmit
  current_month = datepart("m",Date())
  bene_mo_col = (15 + (4*current_month))
  bene_yr_row = 10
  month_count = 0
  abawd_counted_months = 0
  DO
    EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
    IF is_counted_month = "X" or is_counted_month = "M" THEN abawd_counted_months = abawd_counted_months + 1
	bene_mo_col = bene_mo_col - 4
	IF bene_mo_col = 15 THEN
		bene_yr_row = bene_yr_row - 1
		bene_mo_col = 63
	END IF
	month_count = month_count + 1
  LOOP until month_count = 36
  PF3
END function

'And one to case note and end the script. Script will write in each checkbox and the ABAWD status that is built below using previous function and input from worker. 
Function case_note_and_end
	Dialog get_worker_comments
	Cancel_confirmation
	PF3	
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("***Member " & member_number & " has been screened for ABAWD***")
	call write_variable_in_CASE_NOTE(abawd_status)
	IF worked_80_since_closing = 1 AND has_used_second_period <> 1 THEN call write_variable_in_CASE_NOTE("* CL has earned additional 3-month period of ABAWD eligibility.")
	IF worked_80_since_closing = 1 AND has_used_second_period = 1 THEN call write_variable_in_CASE_NOTE("* Client has used 2nd 3 months of eligibility, and 80 hours a month since closure. However they must meet another exemption.")
	IF worked_80_since_closing <> 1 and has_used_second_period = 1 THEN call write_variable_in_CASE_NOTE("* Client has used 2nd 3 months of eligibility, must meet exemption to be eligible for SNAP")
	IF wreg_disa = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are disabled")
	IF care_of_hh_memb = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are responsible for care of a disabled unit member")
	IF age_sixty = 1 THEN call write_variable_in_CASE_NOTE("* Client is over 60.")
	IF under_sixteen = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are under 16.")
	IF sixteen_seventeen = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are age 16 or 17 and living with a parent or caretaker")
	IF care_child_six = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are responsible for the care of a child less than age 6.")
	IF employed_thirty = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are employed 30 hours per week or equivalent to 30 hours a week at minimum wage.") 
	IF unemployment = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are receiving or applied for unemployment insurance.")
	IF enrolled_school = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are enrolled in school/training 1/2 time.")
	IF CD_program = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are enrolled in a sanctioned chemical dependency treatment program.")
	IF receiving_MFIP = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are a MFIP recipient.")
	IF receiving_DWP_WB = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are a DWP/WB recipient.")
	IF applied_SSI = 1 THEN call write_variable_in_CASE_NOTE("* Client states they have applied for SSI.")
	IF age_exempt = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are under 18 or over 50 years old")
	IF cert_preg = 1 THEN call write_variable_in_CASE_NOTE("* Client states certified as pregnant")
	IF working_20 = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are employed 20 hours per week")
	IF dependent_child = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are responsible for the care of a dependent child in the household")
	IF work_exp = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are participating in work experience program")
	IF approved_ET = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are participating in employment and training program")
	IF waiver = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are residing in a waiver area")
	IF receiving_cash = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are a RCA or GA recipient")
	call write_bullet_and_variable_in_CASE_NOTE("Other notes", worker_comment)
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)
	script_end_procedure("")
END Function

'THE SCRIPT=======================================================================================

EMConnect ""

'Checking for maxis, finding case number and getting to blank slate of self. 
call check_for_MAXIS(false)
call MAXIS_case_number_finder(case_number)
back_to_SELF

'Basic info dialog, will reject incorrect case numbers and member numbers
DO
	err_msg = ""
	dialog get_case_number
	cancel_confirmation
    IF case_number = FALSE THEN err_msg = err_msg & vbCr & "Your case number contains characters other than numbers."
    IF len(case_number) > 8 THEN  err_msg = err_msg & vbCr & "Your case number is longer than 8 characters"
    IF len(member_number) = 1 THEN member_number = "0" & member_number  'correcting for 1 digit member numbers
    IF len(member_number) > 2 THEN err_msg = err_msg & vbCr & "Your members number is longer than 2 characters Please use ## format."
    IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
	IF err_msg <> "" THEN MSGBOX err_msg
LOOP until case_number <> "" and worker_signature <> "" and len(member_number) = 2

call check_for_MAXIS(True)

'Logic to check if client is open on GA or RCA as that in itself is an exemption
call navigate_to_screen("stat", "prog")
EMReadScreen cash_I_prog, 2, 6, 67
EMReadScreen cash_I_status, 4, 6, 74
EMReadScreen cash_II_prog, 2, 7, 67
EMReadScreen cash_II_status, 4, 7, 74
IF cash_I_status = "ACTV" and (cash_I_prog = "GA" or cash_I_prog = "RC") THEN script_end_procedure("Client is open on GA or RCA, client is exempt from WREG/ABAWD")
IF cash_II_status = "ACTV" and (cash_II_prog = "GA" or cash_II_prog = "RC") THEN script_end_procedure("Client is open on GA or RCA, client is exempt from WREG/ABAWD")

call navigate_to_screen("stat", "wreg")

'Checking to see if the case is in the county of the worker running it. If it is not the same county then worker cannot case note. 
EMReadScreen User_county_check, 4, 21, 71
EMReadScreen PW_county_check, 4, 21, 21
If User_county_check <> PW_county_check then
	MSGbox "This case is not in your county. You will not be able to case note. A message box will appear at the end of this tool."
	Inquiry_check = "A"
end if

'function to count how many abawd months a specific member has used. 
call how_many_abawd_months(abawd_counted_months)

'Do loop to run dialogs and create ABAWD status variable.
DO
	Dialog wreg_exemptions  'dialog is run asking for input on if client meets any WREG Exmpetions. If they do they are presumed not abawd, if none are checked it continues to next dialog
		cancel_confirmation
		IF under_sixteen = 1 or wreg_disa = 1 or care_of_hh_memb = 1 or age_sixty = 1 or sixteen_seventeen = 1 or care_child_six = 1 or employed_thirty = 1 or unemployment = 1 or enrolled_school = 1 or CD_program = 1 or receiving_MFIP = 1 or receiving_DWP_WB = 1 or applied_SSI = 1 THEN wreg_exempt = true
		IF under_sixteen = 0 and wreg_disa = 0 and care_of_hh_memb = 0 and age_sixty = 0 and sixteen_seventeen = 0 and care_child_six = 0 and employed_thirty = 0 and unemployment = 0 and enrolled_school = 0 and CD_program = 0 and receiving_MFIP = 0 and receiving_DWP_WB = 0 and applied_SSI = 0 THEN wreg_exempt = false
		IF wreg_exempt = TRUE THEN abawd_status = "* Per discussion with client, member " & member_number & " is NOT an ABAWD."
		IF (wreg_exempt = true and PW_county_check <> User_county_check) THEN   'Creating message box for workers screening on out of county cases if member reports WREG exemption.
		  script_end_procedure("Per discussion with client, member " & member_number & " is NOT an ABAWD. Unable to case note due to case being in another county. Process accordingly.")
		Else if wreg_exempt = true THEN
			call case_note_and_end
		  end if
		end if 
  DO
	  Dialog abawd_exemptions  'dialog is run asking for input on if client meets any ABAWD exemptions. IF they do they are presumed not ABAWD, if none are checked it continues to next dialog
		cancel_confirmation
  	    IF waiver = 1 or age_exempt = 1 or cert_preg = 1 or working_20 = 1 or receiving_cash = 1 or dependent_child = 1 or work_exp = 1 or approved_ET = 1 THEN 
		  cl_has_abawd_exemption = true
		  abawd_status = "* Per discussion with client, member " & member_number & " is NOT an ABAWD."
		End If
		IF ButtonPressed = previous_button then exit do
		IF cl_has_abawd_exemption = true and PW_county_check <> User_county_check THEN    'Creating message box for workers screening on out of county cases if member reports ABAWD exemption.
		  script_end_procedure("Per discussion with client, member " & member_number & " is NOT an ABAWD. Unable to case note due to case being in another county. Process accordingly.")
		Else if cl_has_abawd_exemption = true THEN
			call case_note_and_end
		  end if
		end if 
		IF cl_has_abawd_exemption <> true THEN abawd_status = "* Per discussion with client, member " & member_number & " is ABAWD and has used " & abawd_counted_months & " months of SNAP eligibility."
    DO
		Dialog earn_additional_months   'dialog is run asking if client has earned any additional months. If client has not script will case note items checked. 
		cancel_confirmation
		IF ButtonPressed = previous_button then exit do
		IF PW_county_check <> User_county_check THEN    'Creating message box for workers screening on out of county cases if member reports no exemptions.
		  script_end_procedure(abawd_status & chr(13) & "If client has worked at least 80 hours in a month since closing they may be eligible for a 2nd 3 month SNAP period. UNLESS they have already used the 2nd 3 months of eligibility." & chr(13) & "Unable to case note due to case being in another county. Process accordingly.")
		else		
		  call case_note_and_end
		end if
    LOOP until ButtonPressed = -1
  LOOP until ButtonPressed = -1
LOOP until ButtonPressed = -1

script_end_procedure("")
