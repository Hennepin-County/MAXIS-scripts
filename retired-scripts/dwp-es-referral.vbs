'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - DWP ES REFFERAL.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 340                     'manual run time in seconds
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
call changelog_update("01/24/2019", "Updated HIRED in Brooklyn Park phone number per their request.", "MiKayla Handley, Hennepin County")
call changelog_update("03/25/2019", "Updated EMERGE South MPLS orientation dates/times per their request.", "MiKayla Handley, Hennepin County")
call changelog_update("02/06/2019", "Updated EMERGE North MPLS orientation dates/times per their request.", "Ilse Ferris, Hennepin County")
call changelog_update("02/06/2019", "Added HIRED and EMERGE ES providers. Also updated the ES provider droplist to follow the order of orientation locations on the DWP ES CHOICE SHEET.", "Ilse Ferris, Hennepin County")
call changelog_update("12/06/2018", "Brooklyn Center Avivo Thursday sessions discontining effective 12/31/18.", "Ilse Ferris, Hennepin County")
call changelog_update("12/06/2018", "Brooklyn Center Avivo Tuesday sessions discontining effective 12/31/18. Option removed from script as rest of the referral dates fall on holiday days. Also blocked Christmas Day and New Year's day as potential orientation dates.", "Ilse Ferris, Hennepin County")
call changelog_update("12/06/2018", "Bloomington Thrusday sessions discontinued effective 12/14/18. Option removed from script.", "Ilse Ferris, Hennepin County")
call changelog_update("08/03/2018", "Avivo will be closed 08/16/18. The script has been update to exclude this date as a potential orientation date.", "Ilse Ferris, Hennepin County")
call changelog_update("07/20/2018", "Initial version.", "Ilse Ferris, Hennepin County")

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
                appointment_date = dateadd("d", add_days, appointment_date)     'adding the number of days specific to the orientation
            End If
        Next
    Loop until is_holiday = FALSE
End function

'DIALOG----------------------------------------------------------------------------------------------------


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone default screen & 'Searches for a case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
member_number = "01" 'defaults the member_number to 01
'random_date = #3/5/19#

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 291, 130, "ES Letter and Referral"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  EditBox 205, 10, 55, 15, member_number
  DropListBox 60, 35, 125, 15, "Select one..."+chr(9)+"Quick Connect: Northwest"+chr(9)+"Quick Connect: South Mpls"+chr(9)+"Avivo: South MPLS (Tues 1PM)"+chr(9)+"Avivo: South MPLS (Wed 9AM)"+chr(9)+"Avivo: South Subs (Tues 9AM)"+chr(9)+"Avivo: South Subs (Wed 1PM)"+chr(9)+"Avivo: North MPLS (Tues 9AM)"+chr(9)+"Avivo: North MPLS (Wed 9AM)"+chr(9)+"Emerge: South MPLS (Tues 9AM)"+chr(9)+"Emerge: South MPLS (Thurs 9AM)"+chr(9)+"Emerge: North MPLS (Tues 1PM)"+chr(9)+"Emerge: North MPLS (Thurs 9AM)"+chr(9)+"Hired: Brooklyn Park (Mon 1PM)"+chr(9)+"Hired: Brooklyn Park (Wed 9AM)", interview_location
  DropListBox 60, 55, 100, 15, "Select one..."+chr(9)+"Scheduled"+chr(9)+"Rescheduled"+chr(9)+"Rescheduled from WERC", appt_type
  EditBox 205, 55, 75, 15, vendor_num
  EditBox 75, 80, 205, 15, other_referral_notes
  EditBox 75, 105, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 105, 50, 15
    CancelButton 230, 105, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 5, 85, 65, 10, "Other referral notes:"
  Text 150, 15, 55, 10, "HH Memb # (s):"
  Text 10, 110, 60, 10, "Worker Signature:"
  Text 10, 40, 50, 10, "Location/time:"
  Text 20, 60, 35, 10, "Appt type:"
  Text 165, 60, 40, 10, "Vendor #'s:"
EndDialog

'Main dialog
DO
	DO
	    'establishes  that the error message is equal to blank (necessary for the DO LOOP to work)
	    err_msg = ""
	    Dialog Dialog1
		cancel_confirmation   'asks if they really want to cancel script
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF trim(member_number) = "" then err_msg = err_msg & vbNewLine & "* Enter a 2 digit member number, or more than one HH members separated by a comma."
		If interview_location = "Select one..." then err_msg = err_msg & vbNewLine & "* Enter an interview location."
        if appt_type = "Select one..." THEN err_msg = err_msg & vbCr & "* Please enter the appointment type."
        If appt_type <> "Scheduled" and instr(interview_location, "Quick Connect:") THEN err_msg = err_msg & vbCr & "* WERC locations cannot be selected for a reschedule."
		If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'----------------------------------------------------------------------------------------------------ES locations and times
'WERC Northest
IF  interview_location = "Quick Connect: Northwest" THEN
    provider_name = "WERC Northwest Human Service Center"
    provider_address_01 = "7051 Brooklyn Blvd"
    provider_city = "Brooklyn Center"
    provider_ST = "MN"
    provider_zip = "55429"
    provider_phone = "" 'none per project lead's request
    appointment_time = "8:00 - 4:30 PM"
    appointment_date = Date
    provider_row = 13    'WFM1 provider selection based on location
    add_days = 1         'incrementor for DateAdding for orientation options.

'WERC South Mpls
ElseIF interview_location = "Quick Connect: South Mpls" THEN
    provider_name = "WERC South Minneapolis Human Service Center"
    provider_address_01 = "2215 East Lake Street South"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55407"
    provider_phone = "" 'none per project lead's request
    'Date and time
    appointment_time = "8:00 - 4:30 PM"
    appointment_date = Date
    provider_row = 5    'WFM1 provider selection based on location
    add_days = 1         'incrementor for DateAdding for orientation options.

'Avivo South Mpls
ElseIf interview_location = "Avivo: South MPLS (Tues 1PM)" THEN
    provider_name = "Avivo South Mpls"
    provider_address_01 = "900 20th Avenue South"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55404"
    provider_phone = "612-752-8800"
    'Date and time
    appointment_time = "1:00 PM"
    appointment_date = Date + 8 - Weekday(Date, vbTuesday)
    provider_row = 7    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

ElseIf interview_location = "Avivo: South MPLS (Wed 9AM)" THEN
    provider_name = "Avivo South Mpls"
    provider_address_01 = "900 20th Avenue South"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55404"
    provider_phone = "612-752-8800"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbWednesday)
    provider_row = 7    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

'Avivo Bloomington -South Subs
Elseif interview_location = "Avivo: South Subs (Tues 9AM)" then
    provider_name = "Avivo Bloomington"
    provider_address_01 = "2626 East 82nd Street, Suite 370"
    provider_city = "Bloomington"
    provider_ST = "MN"
    provider_zip = "554425"
    provider_phone = "612-752-8942"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbTuesday)
    provider_row = 8    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

Elseif interview_location = "Avivo: South Subs (Wed 1PM)" THEN
    provider_name = "Avivo Bloomington"
    provider_address_01 = "2626 East 82nd Street, Suite 370"
    provider_city = "Bloomington"
    provider_ST = "MN"
    provider_zip = "554425"
    provider_phone = "612-752-8942"
    'Date and time
    appointment_time = "1:00 PM"
    appointment_date = Date + 8 - Weekday(Date, vbWednesday)
    provider_row = 8    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

'Avivo North Mpls
ElseIf interview_location = "Avivo: North MPLS (Tues 9AM)" THEN
    provider_name = "Avivo North Mpls"
    provider_address_01 = "2143 Lowry Avenue North"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55411"
    provider_phone = "612-752-8500"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbTuesday)
    provider_row = 9    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

ElseIf interview_location = "Avivo: North MPLS (Wed 9AM)" THEN
    provider_name = "Avivo North Mpls"
    provider_address_01 = "2143 Lowry Avenue North"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55411"
    provider_phone = "612-752-8500"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbWednesday)
    provider_row = 9    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

'Emerge South MPLS
ElseIf interview_location = "Emerge: South MPLS (Tues 9AM)" THEN
    provider_name = "Emerge South Mpls"
    provider_address_01 = "505 15th Avenue South"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55454"
    provider_phone = "612-876-9301"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbTuesday)
    provider_row = 12     'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

ElseIf interview_location = "Emerge: South MPLS (Thurs 9AM)" THEN
    provider_name = "Emerge South Mpls"
    provider_address_01 = "505 15th Avenue South"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55454"
    provider_phone = "612-876-9301"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbThursday)
    provider_row = 12    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

'Emerge North MPLS
ElseIf interview_location = "Emerge: North MPLS (Tues 1PM)" THEN
    provider_name = "Emerge North Mpls"
    provider_address_01 = "1834 Emerson Avenue North"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55411"
    provider_phone = "612-876-9301"
    'Date and time
    appointment_time = "1:00 PM"
    appointment_date = Date + 8 - Weekday(Date, vbTuesday)
    provider_row = 11    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

ElseIf interview_location = "Emerge: North MPLS (Thurs 9AM)" THEN
    provider_name = "Emerge North Mpls"
    provider_address_01 = "1834 Emerson Avenue North"
    provider_city = "Minneapolis"
    provider_ST = "MN"
    provider_zip = "55411"
    provider_phone = "612-876-9301"
    'Date and time
    appointment_time = "9:00 AM"
    appointment_date = Date + 8 - Weekday(Date, vbThursday)
    provider_row = 11    'WFM1 provider selection based on location
    add_days = 7         'incrementor for DateAdding for orientation options.

'HIRED Brooklyn Park
ElseIf interview_location = "Hired: Brooklyn Park (Mon 1PM)" THEN
   provider_name = "Hired"
   provider_address_01 = "7225 Northland Drive, Suite 100"
   provider_city = "Brooklyn Park"
   provider_ST = "MN"
   provider_zip = "55428"
   provider_phone = "763-210-6215"
   'Date and time
   appointment_time = "1:00 PM"
   appointment_date = Date + 8 - Weekday(Date, vbMonday)
   provider_row = 10   'WFM1 provider selection based on location
   add_days = 7         'incrementor for DateAdding for orientation options.

ElseIf interview_location = "Hired: Brooklyn Park (Wed 9AM)" THEN
   provider_name = "Hired"
   provider_address_01 = "7225 Northland Drive, Suite 100"
   provider_city = "Brooklyn Park"
   provider_ST = "MN"
   provider_zip = "55428"
   provider_phone = "763-210-6215"
   'Date and time
   appointment_time = "9:00 AM"
   appointment_date = Date + 8 - Weekday(Date, vbWednesday)
   provider_row = 10     'WFM1 provider selection based on location
   add_days = 7         'incrementor for DateAdding for orientation options.
End if

'selecting the interview date
DO
	DO
        Call confirm_available_dates
        If appointment_date = random_date then appointment_date = dateadd("d", add_days, appointment_date)            'adding the number of days specific to the orientation
        orientation_date_confirmation = MsgBox("Press YES to confirm the orientation date. For the next week, press NO." & vbNewLine & vbNewLine & _
		"                                                  " & appointment_date & " at " & appointment_time, vbYesNoCancel, "Please confirm the ES orientation referral date")
		If orientation_date_confirmation = vbCancel then script_end_procedure ("The script has ended. An orientation letter has not been sent.")
        If orientation_date_confirmation = vbYes then exit do
        If orientation_date_confirmation = vbNo then appointment_date = dateadd("d", add_days, appointment_date)     'adding the number of days specific to the orientation
	LOOP
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_background_check
member_array = split(member_number, ",")

For each member_number in member_array
    member_number = trim(member_number)
    '----------------------------------------------------------------------------------------------------EMPS: ensures person is valid, collects name and checks for exemption
    Call navigate_to_MAXIS_screen("STAT", "EMPS")
    Call write_value_and_transmit(member_number, 20, 76)
    EMReadScreen memb_error_check, 14, 24, 13
    If memb_error_check = "DOES NOT EXIST" then
        msgbox "The HH member " & member_number & " is invalid. Please review your case if necessary. The script will not continue for this member."
        make_referral = False
    Else
        make_referral = True
        EMReadScreen client_name, 44, 4, 37
        client_name = trim(client_name)

        length = len(client_name)                           'establishing the length of the variable
        position = InStr(client_name, ", ")                  'sets the position at the deliminator (in this case the comma)
        last_name = Left(client_name, position -1)           'establishes client last name as being before the deliminator
        first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator

        first_name = trim(first_name)
        last_name = trim(last_name)
        Client_name = first_name & " " & last_name

        Call fix_case(client_name, 1)
        Call fix_case(first_name, 1)
        client_name = trim(client_name)

        'Ensuring that students have a FSET status of "12" and all others are coded with "30"
        EMReadScreen under_one, 1, 12, 76
        If under_one = "Y" then
        	Do 				'loops until user passwords back in
                exemption_confirmation = MsgBox("Press YES to continue to make the referral. Press NO to skip this member. Press CANCEL to stop the script.", vbYesNoCancel + vbQuestion, "Member is coded as FT care of child under one.")
                IF exemption_confirmation = vbCancel then script_end_procedure("You pressed cancel. The script has ended. No further action taken.")
                IF exemption_confirmation = vbNo then
                    make_referral = False
                    msgbox "You have opted to NOT send a referral for this recepient. Please review your case if necessary. The script will not continue for this member."
                End if
                IF exemption_confirmation = vbCancel then make_referral = True
                CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
            Loop until are_we_passworded_out = false
        End if
    End if

    If make_referral = True then
    	PF9
    	Call create_MAXIS_friendly_date(date, 0, 16, 40)
    	PF3
        back_to_SELF
    END if
Next

For each member_number in member_array
    If make_referral = True then
        'The CASE/NOTE----------------------------------------------------------------------------------------------------
        start_a_blank_CASE_NOTE     'Navigates to a blank case note
        CALL write_variable_in_case_note("**DWP ES referral Appt " & appt_type & " for MEMB " & member_number & "**")
        Call write_variable_in_case_note("* Member referred to ES: #" &  member_number & ", " & client_name)
        CALL write_bullet_and_variable_in_case_note("Appointment date", appointment_date)
        CALL write_bullet_and_variable_in_case_note("Appointment time", appointment_time)
        CALL write_bullet_and_variable_in_case_note("Appointment location", provider_name)
        CALL write_variable_in_case_note("-")
        Call write_bullet_and_variable_in_case_note("Other referral notes", other_referral_notes)
        Call write_bullet_and_variable_in_CASE_NOTE("Vendor #(s)", vendor_num)
        CALL write_variable_in_case_note("---")
        CALL write_variable_in_case_note(worker_signature)

        'Creating verbiaged based on the appt type and location
        If appt_type = "Scheduled" Then
            If instr(interview_location, "Quick Connect:") then
                num_days = "2"
            Else
                num_days = "10"
            End if
        Else
            num_days = "10"
        End if

        Call MAXIS_background_check

        'The SPEC/LETR----------------------------------------------------------------------------------------------------
        Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)    
        Call write_variable_in_SPEC_MEMO("**************DWP ES Orientation Requirement**************")
        Call write_variable_in_SPEC_MEMO("")
        Call write_variable_in_SPEC_MEMO(client_name & " must attend an orientation as part of the Diversionary Work Program (DWP) Employment Services (ES) requirements. Orientation information:")
        Call write_variable_in_SPEC_MEMO("")
        Call write_variable_in_SPEC_MEMO("Date and time: " & appointment_date & " at " & appointment_time)
        Call write_variable_in_SPEC_MEMO(" ")
        Call write_variable_in_SPEC_MEMO("Location: " & provider_name)
        Call write_variable_in_SPEC_MEMO(provider_address_01)
        Call write_variable_in_SPEC_MEMO(provider_city & ", " & provider_ST & " " &  provider_zip)
        If trim(provider_phone) <> "" then Call write_variable_in_SPEC_MEMO("Phone #: " & provider_phone)
        Call write_variable_in_SPEC_MEMO(" ")
        Call write_variable_in_SPEC_MEMO("If " & first_name & "cannot go to the orientation on this date, please contact the career counselor right away. You have " & num_days & " days to make, and keep a new appointment. After 10 days, you will need to return to the county, and reapply for benefits.")
        Call write_variable_in_SPEC_MEMO(" ")
        Call write_variable_in_SPEC_MEMO("If " & first_name & "does not complete this appointment, no benefits will be issued. If " & first_name & "believes he/she has Good Cause for not attending, please contact the career counselor right away.")
        Call write_variable_in_SPEC_MEMO(" ")
        Call write_variable_in_SPEC_MEMO("If " & first_name & "is unable to find suitable childcare, " & first_name & "may bring and keep child/ren with for the entire appointment.")
        PF4		'saves and sends memo
        PF3
        back_to_SELF
    End if
Next

If make_referral = true then
    'Manual referral creation if banked months are used
    Call navigate_to_MAXIS_screen("INFC", "WF1M")			'navigates to WF1M to create the manual referral'
    EMWriteScreen "05", 4, 47								'this is the manual ES staff have identified
    row = 8
    For each member_number in member_array
        IF make_referral = True then
            member_number = trim(member_number)
            EMWriteScreen member_number, row, 9
            EMWriteScreen "DW", row, 46								'enters member number
            Call create_MAXIS_friendly_date(appointment_date, 0, row, 65)			'enters the E & T referral date
            row = row + 1
        End if
    Next

    row = 8
    For each member_number in member_array
        If make_referral = True then
            EmWriteScreen "x", row, 53 'navigates to the ES provider selection screen
            row = row + 1
        End if
    Next
    transmit

    Do
        EMReadScreen ES_popup, 11, 2, 37
        IF ES_popup = "ES Provider" then Call write_value_and_transmit("X", provider_row, 9)
    Loop until ES_popup <> "ES Provider"

    EMWriteScreen appointment_date & ", " & appointment_time & ", " & provider_name, 17, 6		'enters the location, date and time for Hennepin Co ES providers (per request)'
    EmWriteScreen other_referral_notes, 18,6
    PF3
    Call write_value_and_transmit("Y", 11, 64)		'Y to confirm save and saves referral

    script_end_procedure("Your orientation letter, WF1M (manual) referral and case note have been created. Navigate to SPEC/MEMO if you want to review the notice sent to the client.")
else
    script_end_procedure("WF1M referrals we're not made. You may need to review the case and make another referral if necessary.")
End if
