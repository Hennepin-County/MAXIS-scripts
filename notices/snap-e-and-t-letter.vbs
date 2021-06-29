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
call changelog_update("04/19/2021", "Removed Banked Months option for referral. Temporarily removed ABAWD and ABAWD 2nd set for referral option during federal ABAWD waiver.", "Ilse Ferris, Hennepin County")
call changelog_update("02/13/2019", "Added all 2019 county holidays as unavailable orientation dates to select.", "Ilse Ferris, Hennepin County")
call changelog_update("12/17/2018", "Added functionality to block Christmas Day and New Year's day as potential orientation dates.", "Ilse Ferris, Hennepin County")
call changelog_update("12/04/2018", "Updated orientation letter functionality to send SPEC/MEMO. SPEC/LETR retired on 12/01/18.", "Ilse Ferris, Hennepin County")
call changelog_update("02/27/2018", "Added WCOM's regarding voluntary compliance to SPEC/LETR. Updated comments in manual referrals to include 'voluntary' for 30/10 and 30/11 recipients.", "Ilse Ferris, Hennepin County")
call changelog_update("02/27/2018", "Updated to allow referrals for members not coded as mandatory participants under OTHER REFERRAL and WORKING WITH CBO options.", "Ilse Ferris, Hennepin County")
call changelog_update("03/29/2018", "Added ABAWD 2nd set as a referral reason. Removed manual referral option, script will now send a manual referral on all cases. Removed TIKL to follow up on case in 30 days.", "Ilse Ferris, Hennepin County")
call changelog_update("02/27/2018", "Multiple updates include handling for multiple household members, background check, removed exempt counties coding, added other manual reason info into case note, upated TIKL msgbox, and added ABAWD to manual referral droplist. ", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function confirm_available_dates
    Do
        If WeekdayName(WeekDay(appointment_date)) = "Saturday" Then appointment_date = DateAdd("d", 2, appointment_date)
        If WeekdayName(WeekDay(appointment_date)) = "Sunday" Then appointment_date = DateAdd("d", 1, appointment_date)
        is_holiday = FALSE
        For each holiday in HOLIDAYS_ARRAY
            If holiday = appointment_date Then
                is_holiday = TRUE
                appointment_date = dateadd("d", 7, appointment_date)     'adding the number of days specific to the orientation
            End If
        Next
    Loop until is_holiday = FALSE
End function

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone default screen & 'Searches for a case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

member_number = "01"    'defaults the member_number to 01
SNAPET_contact = "the SNAP Employment and Training Team"
SNAPET_phone = "612-596-7411"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 456, 130, "SNAP E&T Appointment Letter"
  EditBox 100, 10, 55, 15, MAXIS_case_number
  EditBox 220, 10, 75, 15, member_number
  DropListBox 100, 35, 195, 15, "Select one..."+chr(9)+"Somali-language (Sabathani, next Tuesday @ 2:00 p.m.)"+chr(9)+"Central NE (HSB, next Wednesday @ 2:00 p.m.)"+chr(9)+"North (HSB, next Wednesday @ 10:00 a.m.)"+chr(9)+"Northwest(Brookdale, next Monday @ 2:00 p.m.)"+chr(9)+"South Mpls (Sabathani, next Tuesday @ 10:00 a.m.)"+chr(9)+"South Suburban (Sabathani, next Tuesday @ 10:00 a.m.)"+chr(9)+"West (Sabathani, next Tuesday @ 10:00 a.m.)", interview_location
  DropListBox 100, 60, 110, 15, "Select one..."+chr(9)+"Other referral"+chr(9)+"Student"+chr(9)+"Working with CBO", manual_referral
  'DropListBox 100, 60, 110, 15, "Select one..."+chr(9)+"ABAWD (3/36 mo.)"+chr(9)+"ABAWD 2nd Set"+chr(9)+"Other referral"+chr(9)+"Student"+chr(9)+"Working with CBO", manual_referral
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
'Main dialog
DO
	DO
	    'establishes  that the error message is equal to blank (necessary for the DO LOOP to work)
	    err_msg = ""
	    Dialog Dialog1
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
        Call confirm_available_dates
        orientation_date_confirmation = MsgBox("Press YES to confirm the orientation date. For the next week, press NO." & vbNewLine & vbNewLine & _
		"                                                  " & appointment_date & " at " & appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & _
		AM_PM, vbYesNoCancel, "Please confirm the SNAP E & T orientation referral date")
		If orientation_date_confirmation = vbCancel then script_end_procedure ("The script has ended. An orientation letter has not been sent.")
		If orientation_date_confirmation = vbYes then exit do
		If orientation_date_confirmation = vbNo then appointment_date = dateadd("d", 7, appointment_date)
	LOOP
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

    first_name = trim(first_name)
    last_name = trim(last_name)
    Call fix_case(first_name, 0)
    Call fix_case(last_name, 0)

    Client_name = trim(first_name) & " " & trim(last_name)

    'Ensuring that the ABAWD_status is "13" for banked months manual referral recipients
    EMReadScreen ABAWD_status, 2, 13, 50
    If manual_referral = "ABAWD 2nd Set" then
        if ABAWD_status <> "11" then script_end_procedure("Member " & member_number & " is not coded as ABAWD 2nd set. The script will now end.")
    Elseif manual_referral = "ABAWD (3/36 mo.)" then
        if ABAWD_status <> "10" then script_end_procedure("Member " & member_number & " is not coded as ABAWD. The script will now end.")
    End if

    'Ensuring that students have a FSET status of "12".
    EMReadScreen FSET_status, 2, 8, 50
    If manual_referral = "Student" and FSET_status <> "12" then script_end_procedure("Member " & member_number & " is not coded as a student. The script will now end.")

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
    start_a_blank_CASE_NOTE         'Navigates to a blank case note
    CALL write_variable_in_case_note("***SNAP E&T Appointment Letter Sent for MEMB " & member_number & "***")
    Call write_variable_in_case_note("* Member referred to E&T: #" &  member_number & ", " & client_name)
    CALL write_bullet_and_variable_in_case_note("Appointment date", appointment_date)
    CALL write_bullet_and_variable_in_case_note("Appointment time", appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM)
    CALL write_bullet_and_variable_in_case_note("Appointment location", SNAPET_name)
    Call write_variable_in_case_note("* The WREG panel has been updated to reflect the E & T orientation date.")
    If manual_referral <> "Select one..." then Call write_variable_in_case_note("* Manual referral made for: " & manual_referral & " recipient.")
    Call write_bullet_and_variable_in_case_note("Other referral notes", other_referral_notes)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)
    PF3

    appointment_info = appointment_date & " at " & appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM & "."
    appointment_location = SNAPET_name & ": " & SNAPET_address_01 & ", " & SNAPET_city & " " & SNAPET_ST & " " & SNAPET_zip

    Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    
    Call write_variable_in_SPEC_MEMO("                     SNAP E&T ORIENTATION")
    Call write_variable_in_SPEC_MEMO(Client_name & " has volunteered to work with SNAP Employment & Training (SNAP E & T).")
    Call write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("Orientation Date:")
    Call write_variable_in_SPEC_MEMO("Your appointment is scheduled for " & appointment_info)
    Call write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("Orientation Location:")
    Call write_variable_in_SPEC_MEMO(appointment_location)
    Call write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("If " & first_name & "cannot attend, contact the SNAP E&T team right away to inquire about additional training dates.")
    Call write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("If " & first_name & "does not attend this orientation, SNAP benefits will remain open if " & first_name & "continues to meet SNAP guidelines. Participation in SNAP E&T services is voluntary, but participation in these services could help you:")
    Call write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("* Develop skills to find suitable employment")
    Call write_variable_in_SPEC_MEMO("* Offer training and supportive services")
    Call write_variable_in_SPEC_MEMO("* Provide help with transportation and other work-related costs")
    Call write_variable_in_SPEC_MEMO("* Qualify for additional SNAP benefits")
    Call write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("Your benefits will not be affected if you do not participate - the choice is yours. Please contact the SNAP Employment and Training team at 612-596-7411 with any additional questions. Thank you.")

    PF4 'to save
    PF3 'to exit
Next

'Manual referral creation if banked months are used
Call navigate_to_MAXIS_screen("INFC", "WF1M")			'navigates to WF1M to create the manual referral'
EMWriteScreen "99", 4, 47								'this is the manual referral code that DHS has approved
IF manual_referral = "Student" then
	EMWriteScreen "Student", 17, 6
ELSEIF manual_referral = "Working with CBO" then
	EMWriteScreen "Working with Community Based Organization", 17, 6
ELSEIF manual_referral = "Other referral" then
	EMWriteScreen other_referral_notes, 17, 6
ELSEIF manual_referral = "ABAWD (3/36 mo.)" then
	EMWriteScreen "ABAWD (3/36 mo.) - Voluntary", 17, 6
ELSEIF manual_referral = "ABAWD 2nd Set" then
	EMWriteScreen "ABAWD 2nd Set - Voluntary", 17, 6
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

script_end_procedure("Your orientation letter, WF1M (manual) referral and case note have been created. Navigate to SPEC/MEMO if you want to review the notice sent to the client.")
'vbNewLine & vbNewLine & "Please ensure that you have sent the form ""ABAWD FS RULES"" to the client.")
