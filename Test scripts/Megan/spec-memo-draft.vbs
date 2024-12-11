'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - PROJECT NOOB SCRIPT.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone


'MODIFIED create_appointment_letter_notice_recertification(programs, intvw_programs, interview_end_date, last_day_of_recert)
function create_appointment_letter(application_date, interview_date, last_contact_day)
'--- This function standardizes the creation and content for an application appointment letter notice.
'~~~~~ application_date: Date of applicattion, must be in date format
'~~~~~ interview_date: Date interview was conducted, must be in date format
'~~~~~ last_contact_day: Date of last contact with resident, must be in date format
'===== Keywords: MAXIS, ADMIN, NOTICE, Appointment, Application
  Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
  Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
  Call write_variable_in_SPEC_MEMO(" ")
  Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & interview_date & ". **")
  Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
  Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 3:00pm Monday thru Friday.")
  Call write_variable_in_SPEC_MEMO(" ")
  Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
  Call write_variable_in_SPEC_MEMO(" ")
  Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
  Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
  Call write_variable_in_SPEC_MEMO(" ")
  CALL write_variable_in_SPEC_MEMO("You can complete interviews by phone. In-person support is available at several service center locations(M-F 8-4:30)")
  CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
  CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
  CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
  CALL write_variable_in_SPEC_MEMO(" - 525  Portland Ave S (5th floor), Minneapolis 55415")
  CALL write_variable_in_SPEC_MEMO(" - 9600 Aldrich Ave S, Bloomington 55420")
  Call write_variable_in_SPEC_MEMO(" ")
  CALL digital_experience
  Call write_variable_in_SPEC_MEMO(" ")
  CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can always request a paper copy via phone.")
  PF4
end function

'MODIFIED create_appointment_letter_notice_application(application_date, interview_date, last_contact_day)
function create_appointment_letter_notice_recert(programs, intvw_programs, interview_end_date, last_day_of_recert)
'--- This function standardizes the creation and content for a recertification appointment letter notice
'~~~~~ programs: The variable to identify the programs applicable to the case
'~~~~~ intvw_programs: The variable to identify the program identified during the interview (MFIP/SNAP, MFIP, SNAP)
'~~~~~ interview_end_date: Last date an interview can be conducted within the recertification window, must be in date format
'~~~~~ last_day_of_recert: Last day of recertification, must be in date format
'===== Keywords: MAXIS, ADMIN, NOTICE, Appointment, Recertification
	CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
	If len(programs) < 11 Then CALL write_variable_in_SPEC_MEMO("")
	CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & intvw_programs & " case to continue.")
	CALL write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("  *** Please complete your interview by " & interview_end_date & ". ***")
	Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
	Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 3:00pm Monday thru Friday.")
	CALL write_variable_in_SPEC_MEMO("")
	If len(programs) < 11 Then
		CALL write_variable_in_SPEC_MEMO("**  Your " & programs & " case will close on " & last_day_of_recert & " unless  **")
	ElseIf len(programs) > 14 Then
		CALL write_variable_in_SPEC_MEMO("*Your " & programs & " case will close on " & last_day_of_recert & " unless")
	ElseIf len(programs) > 10 Then
		CALL write_variable_in_SPEC_MEMO("* Your " & programs & " case will close on " & last_day_of_recert & " unless *")
	End If
	CALL write_variable_in_SPEC_MEMO("** we receive your paperwork and complete the interview. **")
	CALL write_variable_in_SPEC_MEMO("")
	CALL write_variable_in_SPEC_MEMO("You can complete interviews by phone. In-person support is available at several service center locations(M-F 8-4:30)")
	CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
	CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
	CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
	CALL write_variable_in_SPEC_MEMO(" - 525  Portland Ave S (5th floor), Minneapolis 55415")
	CALL write_variable_in_SPEC_MEMO(" - 9600 Aldrich Ave S, Bloomington 55420")
	CALL write_variable_in_SPEC_MEMO(" ")
	Call digital_experience
	CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can always request a paper copy via phone.")
	PF4         'Submit the MEMO
end function

'MODIFIED create_NOMI_application(application_date, appt_date, last_contact_day)
function create_NOMI_appl(application_date, appt_date, last_contact_day)
'--- This function standardizes the creation and content for a NOMI (Notice of Missed Interview) application notice.
'~~~~~ application_date: Date CAF was received, must be in date format
'~~~~~ appt_date: Date interview is to be conducted by, must be in date format. This is 7 days from the CAF date.
'~~~~~ last_contact_day: Date resident has to contact the county, must be in date format. This is 30 days from the application date.
	Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
	Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & appt_date & ".")
	Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
	Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 3:00pm Monday thru Friday.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
	Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
	Call write_variable_in_SPEC_MEMO(" ")
	CALL write_variable_in_SPEC_MEMO("You can complete interviews by phone. In-person support is available at several service center locations(M-F 8-4:30)")
	CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
	CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
	CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
	CALL write_variable_in_SPEC_MEMO(" - 525  Portland Ave S (5th floor), Minneapolis 55415")
	CALL write_variable_in_SPEC_MEMO(" - 9600 Aldrich Ave S, Bloomington 55420")
	CALL write_variable_in_SPEC_MEMO(" More detail can be found at hennepin.us/economic-supports")
	CALL write_variable_in_SPEC_MEMO("")
	CALL digital_experience
	PF4
end function

' MODIFED create_NOMI_recertification(caf_date_as_of_today, last_day_of_recert)
function create_NOMI_recert(caf_date_as_of_today, last_day_of_recert)
'--- This function standardizes the creation and content for a NOMI (Notice of Missed Interview) recertification notice.
'~~~~~ caf_date_as_of_today: Date recertification paperwork was received if received, must be in date format.
'~~~~~ last_day_of_recert: Last day to complete an interview in order to continue benefits, must be in date format.
	if caf_date_as_of_today <> "" then CALL write_variable_in_SPEC_MEMO("We received your Recertification Paperwork on " & caf_date_as_of_today & ".")
	if caf_date_as_of_today = "" then CALL write_variable_in_SPEC_MEMO("Your Recertification Paperwork has not yet been received.")
	CALL write_variable_in_SPEC_MEMO("")
	CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
	CALL write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
	Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 3:00pm Monday thru Friday.")
	CALL write_variable_in_SPEC_MEMO("")
	CALL write_variable_in_SPEC_MEMO("You can complete interviews by phone. In-person support is available at several service center locations(M-F 8-4:30)")
	CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
	CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
	CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
	CALL write_variable_in_SPEC_MEMO(" - 525  Portland Ave S (5th floor), Minneapolis 55415")
	CALL write_variable_in_SPEC_MEMO(" - 9600 Aldrich Ave S, Bloomington 55420")
	CALL write_variable_in_SPEC_MEMO(" More detail can be found at hennepin.us/economic-supports")
	CALL write_variable_in_SPEC_MEMO("")
	CALL digital_experience
	CALL write_variable_in_SPEC_MEMO(" ")
	CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
	CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")
	PF4         'Submit the MEMO
end function



'Hardcoding values for testing
programs = "MFIP, SNAP, CASH"
intvw_programs = "SNAP, CASH, MFI"
interview_end_date = "05/05/2024"
caf_date_as_of_today = "03/03/2024"
last_day_of_recert = "04/04/2024"
application_date = "03/03/2024"
appt_date = "03/03/2024"
interview_date = "03/24/2024"
last_contact_day = "04/04/2024"
MAXIS_case_number = "333036"
MAXIS_footer_month = "04"
MAXIS_footer_year = "24"

'DIALOG
BeginDialog Dialog1, 0, 0, 191, 105, "NOOB Test Case"
Text 5, 10, 50, 10, "Case Number:"
EditBox 75, 5, 45, 15, MAXIS_case_number
Text 5, 30, 65, 10, "Footer Month/year:"
EditBox 75, 25, 20, 15, MAXIS_footer_month
EditBox 100, 25, 20, 15, MAXIS_footer_year
ButtonGroup ButtonPressed
    OkButton 75, 85, 50, 15
    CancelButton 135, 85, 50, 15
EndDialog
   
DO
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
        IF MAXIS_case_number = "" or (IsNumeric(MAXIS_case_number) = False) or (LEN(MAXIS_case_number) > 8) Then err_msg = "Case Number: Must have numeric entry <8 characters" & vbNewLine             
        IF MAXIS_footer_month = "" or (IsNumeric(MAXIS_footer_month) = False) or (LEN(MAXIS_footer_month) <> 2) Then err_msg = err_msg & vbNewLine & "Month: 2 Characters and numeric" & vbNewLine
        IF MAXIS_footer_year = "" or (IsNumeric(MAXIS_footer_year) = False) or (LEN(MAXIS_FOOTER_year) <> 2) Then err_msg = err_msg & vbNewLine & "Year: 2 Characters & numeric" & vbNewLine
        If err_msg <> "" Then Msgbox "***Notice***" & vbNewLine & err_msg 
        'Add in all of your mandatory field handling from your dialog here. Does not restrict user to 2 or 8 digits....gap
            'Add to all dialogs where you need to work within BLUEZONE
    Loop Until err_msg = ""
    CALL check_for_password(are_we_passworded_out)          'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false                    'loops until user passwords back in
           
'CALL SPECIFIC spec/memo
call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True) 

'MODIFIED NOTICES
' Call create_appointment_letter_notice_recert(programs, intvw_programs, interview_end_date, last_day_of_recert)
' Call create_appointment_letter(application_date, interview_date, last_contact_day)
' Call create_NOMI_appl(application_date, appt_date, last_contact_day)
' Call create_NOMI_recert(caf_date_as_of_today, last_day_of_recert)

'ORIGINAL NOTICES
' Call create_appointment_letter_notice_recertification(programs, intvw_programs, interview_end_date, last_day_of_recert)
' Call create_appointment_letter_notice_application(application_date, interview_date, last_contact_day)
' Call create_NOMI_application(application_date, appt_date, last_contact_day)
' Call create_NOMI_recertification(caf_date_as_of_today, last_day_of_recert)