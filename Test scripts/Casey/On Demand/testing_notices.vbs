'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - RECERTIFICATIONS.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 304			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================
'run_locally = TRUE
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================
EMConnect ""

Call MAXIS_case_number_finder (MAXIS_case_number)

BeginDialog Dialog1, 0, 0, 191, 50, "Dialog"
  DropListBox 10, 30, 100, 45, "Pick"+chr(9)+"RECERT - APPT Notice"+chr(9)+"RECERT - NOMI"+chr(9)+"RECERT - VERIFS"+chr(9)+"APPLICATION - APPT Notice"+chr(9)+"APPLICATION - NOMI", memo_to_send
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  EditBox 60, 5, 50, 15, MAXIS_case_number
EndDialog

Do
    err_msg = ""

    dialog Dialog1
    if buttonpressed = 0 Then stopscript
    if IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "Invalid MAXIS Case Number"
    if len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "Invalid MAXIS Case Number"
    If memo_to_send = "Pick" Then err_msg = err_msg & vbNewLine & "Pick Notice"

    If err_msg <> "" Then MsgBox "Fix:" & vbNewLine & err_msg
Loop until err_msg = ""


Call start_a_new_spec_memo

If memo_to_send = "RECERT - APPT Notice" Then
    'OD Recertifications - APPOINTMENT NOTICE

    programs = "MFIP/SNAP"
    last_day_of_recert = CM_plus_1_mo & "/30/" & CM_plus_1_yr
    interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr
    'NOTICE ON LINE 768'
    'EMSendKey("************************************************************")           'for some reason this is more stable then using write_variable
    CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
    CALL write_variable_in_SPEC_MEMO("")

    CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & programs & " case to continue.")

    ' CALL write_variable_in_SPEC_MEMO("** Your " & programs & " case will close on " & last_day_of_recert & " **")
    ' CALL write_variable_in_SPEC_MEMO("if we do not complete an interview.")
    ''"To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday. Please complete your interview by " & interview_end_date & ".")
    CALL write_variable_in_SPEC_MEMO("")

    Call write_variable_in_SPEC_MEMO("  *** Please complete your interview by " & interview_end_date & ". ***")
    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    'Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
    Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday thru Friday.")
    CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ".")
    ' 'CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("You must also complete an interview for your " & programs & " case to continue.")
    CALL write_variable_in_SPEC_MEMO("**  Your " & programs & " case will close on " & last_day_of_recert & " unless    **")
    CALL write_variable_in_SPEC_MEMO("** we receive your paperwork and complete the interview. **")
    CALL write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")

    Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    Call write_variable_in_SPEC_MEMO(" ")


    CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")




    'Call write_variable_in_SPEC_MEMO("Some cases are eligible to have SNAP benefits issued within 24 hours of the interview, call right away if you have an urgent need.")
    'Call write_variable_in_SPEC_MEMO("Interviews can also be completed in person at one of our six offices:")

    ' Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
    ' Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    '
    '
    '
    ' CALL write_variable_in_SPEC_MEMO("We must have your renewal paperwork to do your interview. Please send proofs with your renewal paperwork.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house payment receipt, mortgage, lease, subsidy, etc.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed): prescription and medical bills, etc.")
    ' CALL write_variable_in_SPEC_MEMO("")

ElseIf memo_to_send = "RECERT - VERIFS" Then

    CALL write_variable_in_SPEC_MEMO("As a part of the Renewal Process we must receive recent verification of your inforomation. To speed the renewl process, please send proofs with your renewal paperwork.")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, employer statement,")
    CALL write_variable_in_SPEC_MEMO("   income reports, business ledgers, income tax forms, etc.")
    CALL write_variable_in_SPEC_MEMO("   *If a job has ended, send proof of the end of employment")
    CALL write_variable_in_SPEC_MEMO("   and last pay.")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house")
    CALL write_variable_in_SPEC_MEMO("   payment receipt, mortgage, lease, subsidy, etc.")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed):")
    CALL write_variable_in_SPEC_MEMO("   prescription and medical bills, etc.")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

ElseIf memo_to_send = "RECERT - NOMI" Then
    'OD Recertifications - NOMI
    recvd_appl = FALSE
    date_of_app = DateAdd("d", -5, date)
    last_day_of_recert = CM_mo & "/30/" & CM_yr

    'NOTICE ON LINE 902'
    if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("We received your Recertification Paperwork on " & date_of_app & ".")
    if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Your Recertification Paperwork has not yet been received.")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
    CALL write_variable_in_SPEC_MEMO("")


    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday thru Friday.")
    CALL write_variable_in_SPEC_MEMO("")
    Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
    Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")


    ' CALL write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("You may also come in to the office to complete an interview between 8:00 am and 4:30pm Monday through Friday.")
    CALL write_variable_in_SPEC_MEMO("")
    CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
    CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")


ElseIf memo_to_send = "APPLICATION - APPT Notice" Then  'THIS ONE IS DONE AND ILSE IS VETTING'
    'OD Application - APPOINTMENT NOTICE
    application_date = DateAdd("d", -2, date)
    need_intv_date = DateAdd("d", 5, date)
    last_contact_day = DateAdd("d", 28, date)

    'NOTICE ON LINE 1113'
    'EMsendkey("************************************************************")
    'Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & application_date & ".")
    Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
    Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & need_intv_date & ". **")
    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    'Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
    Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday thru Friday.")
    'Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Mon through Fri.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    'Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday through Friday.")
    Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")

    'Call write_variable_in_SPEC_MEMO("Some cases are eligible to have SNAP benefits issued within 24 hours of the interview, call right away if you have an urgent need.")
    'Call write_variable_in_SPEC_MEMO("Interviews can also be completed in person at one of our six offices:")
    Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
    Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")
    'Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3).")
    'Call write_variable_in_SPEC_MEMO("************************************************************")

ElseIf memo_to_send = "APPLICATION - NOMI" Then
    'OD Application - NOMI
    application_date = DateAdd("d", -7, date)
    appointment_date = date
    nomi_last_contact_day = DateAdd("d", 23, date)

    'NOTICE ON LINE 1223'
    'EMsendkey("************************************************************")
    Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
    Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & appointment_date & ".")
    Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")


    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    'Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
    Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday thru Friday.")
    'Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Mon through Fri.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    'Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday through Friday.")
    Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")

    'Call write_variable_in_SPEC_MEMO("Some cases are eligible to have SNAP benefits issued within 24 hours of the interview, call right away if you have an urgent need.")
    'Call write_variable_in_SPEC_MEMO("Interviews can also be completed in person at one of our six offices:")
    Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & nomi_last_contact_day & " **")
    Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
    Call write_variable_in_SPEC_MEMO(" ")
    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")


    '
    ' Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at ")
    ' Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("If you do not complete the interview by " & nomi_last_contact_day & " your application will be denied.") 'add 30 days
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to- face interview.")
    ' Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    ' Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
    ' Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
    ' Call write_variable_in_SPEC_MEMO("************************************************************")
End If

script_end_procedure("")
