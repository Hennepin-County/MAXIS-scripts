'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - SNAP E AND T LETTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 280                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("03/29/2018", "Added ABAWD 2nd set as a referral reason. Removed manual referral option, script will now send a manual referral on all cases. Removed TIKL to follow up on case in 30 days.", "Ilse Ferris, Hennepin County")
call changelog_update("02/27/2018", "Multiple updates include handling for multiple household members, background check, removed exempt counties coding, added other manual reason info into case note, upated TIKL msgbox, and added ABAWD to manual referral droplist. ", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG----------------------------------------------------------------------------------------------------
'This is a Hennepin specific dialog, should not be used for other counties!!!!!!!!
BeginDialog SNAPET_Hennepin_dialog, 0, 0, 456, 130, "SNAP E&T Appointment Letter"
  EditBox 100, 10, 55, 15, MAXIS_case_number
  EditBox 220, 10, 75, 15, member_number
  DropListBox 100, 35, 195, 15, "Select one..."+chr(9)+"Somali-language (Sabathani, next Tuesday @ 2:00 p.m.)"+chr(9)+"Central NE (HSB, next Wednesday @ 2:00 p.m.)"+chr(9)+"North (HSB, next Wednesday @ 10:00 a.m.)"+chr(9)+"Northwest(Brookdale, next Monday @ 2:00 p.m.)"+chr(9)+"South Mpls (Sabathani, next Tuesday @ 10:00 a.m.)"+chr(9)+"South Suburban (Sabathani, next Tuesday @ 10:00 a.m.)"+chr(9)+"West (Sabathani, next Tuesday @ 10:00 a.m.)", interview_location
  DropListBox 100, 60, 110, 15, "Select one..."+chr(9)+"ABAWD (3/36 mo.)"+chr(9)+"ABAWD 2nd Set"+chr(9)+"Banked months"+chr(9)+"Other referral"+chr(9)+"Student"+chr(9)+"Working with CBO", manual_referral
  EditBox 100, 80, 195, 15, other_referral_notes
  EditBox 100, 105, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 105, 50, 15
    CancelButton 245, 105, 50, 15
  Text 5, 40, 95, 10, "Client's region of residence: "
  Text 40, 65, 55, 10, "Referral reason:"
  GroupBox 305, 10, 145, 115, "For non-English speaking ABAWD's:"
  Text 45, 15, 50, 10, "Case Number:"
  Text 25, 85, 70, 10, "Other referral reason:"
  Text 165, 15, 55, 10, "HH Memb # (s):"
  Text 35, 110, 60, 10, "Worker Signature:"
  Text 315, 25, 130, 35, "If your client is requsting a Somali-language orientation, select this option in the 'client's region of residence' field."
  Text 315, 65, 130, 55, "For all other languages, do not use this script. Contact E and T staff, and request language-specific SNAP E and T Orientation/intake. Provide client with the E and T contact information, and instruct them to contact them to schedule orientation within one week."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone default screen & 'Searches for a case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'defaults the member_number to 01
member_number = "01"
SNAPET_contact = "the SNAP Employment and Training Team"
SNAPET_phone = "612-596-7411"

'Main dialog
DO
	DO
	    'establishes  that the error message is equal to blank (necessary for the DO LOOP to work)
	    err_msg = ""
	    Dialog SNAPET_Hennepin_dialog
	    'CO #27 HENNEPIN COUNTY addresses, date and times of orientations
	    'Somali-language orientation
	    IF interview_location = "Somali-language (Sabathani, next Tuesday @ 2:00 p.m.)" then
	    	SNAPET_name = "Sabathani Community Center"
	    	SNAPET_address_01 = "310 East 38th Street #120"
	    	SNAPET_city = "Minneapolis"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55409"
	    	appointment_time_prefix_editbox = "02"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "PM"
	    	appointment_date = Date + 8 - Weekday(Date, vbTuesday)
	    'Central NE
	    Elseif interview_location = "Central NE (HSB, next Wednesday @ 2:00 p.m.)" THEN
	    	SNAPET_name = "Health Services Building"
	    	SNAPET_address_01 = "525 Portland Ave, 5th floor"
	    	SNAPET_city = "Minneapolis"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55415"
	    	appointment_time_prefix_editbox = "02"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "PM"
	    	appointment_date = Date + 8 - Weekday(Date, vbWednesday)
	    'North
	    ElseIF interview_location = "North (HSB, next Wednesday @ 10:00 a.m.)" THEN
	    	SNAPET_name = "Health Services Building"
	    	SNAPET_address_01 = "525 Portland Ave, 5th floor"
	    	SNAPET_city = "Minneapolis"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55415"
	    	appointment_time_prefix_editbox = "10"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "AM"
	    appointment_date = Date + 8 - Weekday(Date, vbWednesday)
	    'Northwest
	    ElseIf interview_location = "Northwest(Brookdale, next Monday @ 2:00 p.m.)" THEN
	    	SNAPET_name = "Brookdale Human Services Center"
	    	SNAPET_address_01 = "6125 Shingle Creek Parkway, Suite 400"
	    	SNAPET_city = "Brooklyn Center"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55430"
	    	appointment_time_prefix_editbox = "02"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "PM"
	    	appointment_date = Date + 8 - Weekday(Date, vbMonday)
	    'South Minneapolis
	    ElseIf interview_location = "South Mpls (Sabathani, next Tuesday @ 10:00 a.m.)" THEN
	    	SNAPET_name = "Sabathani Community Center"
	    	SNAPET_address_01 = "310 East 38th Street #120"
	    	SNAPET_city = "Minneapolis"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55409"
	    	appointment_time_prefix_editbox = "10"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "AM"
	    	appointment_date = Date + 8 - Weekday(Date, vbTuesday)
	    'South Suburban
	    ElseIf interview_location = "South Suburban (Sabathani, next Tuesday @ 10:00 a.m.)" THEN
	    	SNAPET_name = "Sabathani Community Center"
	    	SNAPET_address_01 = "310 East 38th Street #120"
	    	SNAPET_city = "Minneapolis"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55409"
	    	appointment_time_prefix_editbox = "10"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "AM"
	    	appointment_date = Date + 8 - Weekday(Date, vbTuesday)
	    'West
	    ElseIf interview_location = "West (Sabathani, next Tuesday @ 10:00 a.m.)" THEN
	    	SNAPET_name = "Sabathani Community Center"
	    	SNAPET_address_01 = "310 East 38th Street #120"
	    	SNAPET_city = "Minneapolis"
	    	SNAPET_ST = "MN"
	    	SNAPET_zip = "55409"
	    	appointment_time_prefix_editbox = "10"
	    	appointment_time_post_editbox = "00"
	    	AM_PM = "AM"
	    	appointment_date = Date + 8 - Weekday(Date, vbTuesday)
	    END IF
	
		'asks if they really want to cancel script
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF trim(member_number) = "" then err_msg = err_msg & vbNewLine & "* Enter a 2 digit member number, or more than one HH members separated by a comma."
		If interview_location = "Select one..." then err_msg = err_msg & vbNewLine & "* Enter an interview location."
        If manual_referral = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a referral reason."
		IF (manual_referral = "Other referral" and other_referral_notes = "") then err_msg = err_msg & vbNewLine & "* Enter other manual referral notes."
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'selecting the interview date 
DO
	DO
		orientation_date_confirmation = MsgBox("Press YES to confirm the orientation date. For the next week, press NO." & vbNewLine & vbNewLine & _
		"                                                  " & appointment_date & " at " & appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & _
		AM_PM, vbYesNoCancel, "Please confirm the SNAP E & T orientation referral date")
		If orientation_date_confirmation = vbCancel then script_end_procedure ("The script has ended. An orientation letter has not been sent.")
		If orientation_date_confirmation = vbYes then exit do
		If orientation_date_confirmation = vbNo then appointment_date = dateadd("d", 7, appointment_date)
	LOOP until orientation_date_confirmation = vbYes
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_background_check
member_array = split(member_number, ",")

For each member_number in member_array 
    member_number = trim(member_number)
    'Updates the WREG panel with the appointment_date
    Call navigate_to_MAXIS_screen("STAT", "WREG")
    Call write_value_and_transmit(member_number, 20, 76)
    EMReadScreen memb_error_check, 20, 24, 15 
    If memb_error_check = "NOT IN THE HOUSEHOLD" then script_end_procedure ("The HH member " & member_number & " is invalid. Please review your case if necessary. The script will not continue for this member.")
    
    EMReadScreen client_name, 44, 4, 37
    client_name = trim(client_name)
    
    if instr(client_name, ", ") then    						'Most cases have both last name and 1st name. This seperates the two names
    	length = len(client_name)                           'establishing the length of the variable
    	position = InStr(client_name, ", ")                  'sets the position at the deliminator (in this case the comma)
    	last_name = Left(client_name, position -1)           'establishes client last name as being before the deliminator
        first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
        If instr(first_name, " ") then first_name = left(first_name, len(first_name) - 2)
    END IF

    Client_name = first_name & " " & last_name
        
    'Ensuring that students have a FSET status of "12" and all others are coded with "30"
    EMReadScreen FSET_status, 2, 8, 50
    If manual_referral <> "Other referral" or manual_referral <> "Working with CBO" then  
        If manual_referral = "Student" then
            if FSET_status <> "12" then script_end_procedure("Member " & member_number & " is not coded as a student. The script will now end.")
        Else 
            If FSET_status <> "30" then script_end_procedure("Member " & member_number & " is not coded as a Mandatory FSET Participant. The script will now end.")
        End if
        
        'Ensuring that the ABAWD_status is "13" for banked months manual referral recipients
        EMReadScreen ABAWD_status, 2, 13, 50
        If manual_referral = "Banked months" then
         	if ABAWD_status <> "13" then script_end_procedure("Member " & member_number & " is not coded as a banked months recipient. The script will now end.")
        Elseif manual_referral = "ABAWD 2nd Set" then
            if ABAWD_status <> "11" then script_end_procedure("Member " & member_number & " is not coded as ABAWD 2nd set. The script will now end.")
        Elseif manual_referral = "ABAWD (3/36 mo.)" then
            if ABAWD_status <> "10" then script_end_procedure("Member " & member_number & " is not coded as ABAWD. The script will now end.")
        End if
    End if 
    
    'Ensuring the orientation date is coding in the with the referral date scheduled
    EMReadScreen orientation_date, 8, 9, 50
    orientation_date = replace(orientation_date, " ", "/")
    If appointment_date <> orientation_date then
    	PF9
    	Call create_MAXIS_friendly_date(appointment_date, 0, 9, 50)
    	PF3
    END if
Next 
        
For each member_number in member_array 
    'The CASE/NOTE----------------------------------------------------------------------------------------------------
    'Navigates to a blank case note
    start_a_blank_CASE_NOTE
    CALL write_variable_in_case_note("***SNAP E&T Appointment Letter Sent for MEMB " & member_number & " ***")
    Call write_variable_in_case_note("* Member referred to E&T: #" &  member_number & ", " & client_name)
    CALL write_bullet_and_variable_in_case_note("Appointment date", appointment_date)
    CALL write_bullet_and_variable_in_case_note("Appointment time", appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM)
    CALL write_bullet_and_variable_in_case_note("Appointment location", SNAPET_name)
    Call write_variable_in_case_note("* The WREG panel has been updated to reflect the E & T orientation date.")
    If manual_referral <> "Select one..." then Call write_variable_in_case_note("* Manual referral made for: " & manual_referral & " recipient.")
    Call write_bullet_and_variable_in_case_note("Other referral notes", other_referral_notes)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)
    
    'The SPEC/LETR----------------------------------------------------------------------------------------------------
    call navigate_to_MAXIS_screen("SPEC", "LETR")
    'Opens up the SNAP E&T Orientation LETR. If it's unable the script will stop.
    EMWriteScreen "x", 8, 12
    transmit
    EMReadScreen LETR_check, 4, 2, 49
    If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
    
    'Writes the info into the LETR.
    IF len(appointment_time_prefix_editbox) = 1 THEN appointment_time_prefix_editbox = "0" & appointment_time_prefix_editbox 'This prevents the letter from being cancelled due to single digit hour
    EMWriteScreen client_name, 4, 28
    call create_MAXIS_friendly_date_three_spaces_between(appointment_date, 0, 6, 28)
    EMWriteScreen appointment_time_prefix_editbox, 7, 28
    EMWriteScreen appointment_time_post_editbox, 7, 33
    EMWriteScreen AM_PM, 7, 38
    EMWriteScreen SNAPET_name, 9, 28
    EMWriteScreen SNAPET_address_01, 10, 28
    EMWriteScreen SNAPET_city & ", " & SNAPET_ST & " " &  SNAPET_zip, 11, 28
    call create_MAXIS_friendly_phone_number(SNAPET_phone, 13, 28) 'takes out non-digits if listed in variable, and formats phone number for the field
    EMWriteScreen SNAPET_contact, 16, 28
    PF4		'saves and sends memo
    PF3
    PF3
Next 

'Manual referral creation if banked months are used
Call navigate_to_MAXIS_screen("INFC", "WF1M")			'navigates to WF1M to create the manual referral'
EMWriteScreen "99", 4, 47								'this is the manual referral code that DHS has approved
												'this is a program for ABAWD's for SNAP is the only option for banked months
If manual_referral = "Banked months" then
	EMWriteScreen "Banked ABAWD month referral, initial month", 17, 6	'DHS wants these referrals marked, this marks them
ELSEIF manual_referral = "Student" then
	EMWriteScreen "Student", 17, 6
ELSEIF manual_referral = "Working with CBO" then
	EMWriteScreen "Working with Community Based Organization", 17, 6
ELSEIF manual_referral = "Other referral" then
	EMWriteScreen other_referral_notes, 17, 6
ELSEIF manual_referral = "ABAWD (3/36 mo.)" then
	EMWriteScreen "ABAWD (3/36 mo.)", 17, 6    
ELSEIF manual_referral = "ABAWD 2nd Set" then
	EMWriteScreen "ABAWD 2nd Set", 17, 6  
END IF

row = 8
For each member_number in member_array
    member_number = trim(member_number)
    EMWriteScreen member_number, row, 9		
    EMWriteScreen "FS", row, 46								'enters member number
    Call create_MAXIS_friendly_date(appointment_date, 0, row, 65)			'enters the E & T referral date
    row = row + 1
Next 
																																
row = 8
For each member_number in member_array
    EmWriteScreen "x", row, 53 'navigates to the ES provider selection screen
    row = row + 1
Next
transmit
Do 
    EMReadScreen ES_popup, 11, 2, 37
    IF ES_popup = "ES Provider" then Call write_value_and_transmit("X", 5, 9)
Loop until ES_popup <> "ES Provider"
    												
EMWriteScreen appointment_date & ", " & appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM & ", " & SNAPET_name, 18, 6		'enters the location, date and time for Hennepin Co ES providers (per request)'
PF3			
Call write_value_and_transmit("Y", 11, 64)		'Y to confirm save and saves referral

script_end_procedure("Your orientation letter, WF1M (manual) referral and case note have been created. Navigate to SPEC/WCOM if you want to review the notice sent to the client." & _
vbNewLine & vbNewLine & "Make sure that you have made your E & T referral, and that you have sent the form ""ABAWD FS RULES"" to the client.")
