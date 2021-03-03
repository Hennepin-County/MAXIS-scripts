'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - Individual Recert.vbs"
start_time = timer
STATS_counter = 1			 'sets the stats counter at one
STATS_manualtime = 180			 'manual run time in seconds
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
call changelog_update("03/02/2021", "Update EZ Info Phone hours from 9-4 pm to 8-4:30 pm.", "Ilse Ferris, Hennepin County")
call changelog_update("10/9/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'ADDITIONAL FUNCTIONS ======================================================================================================
function convert_date_to_day_first(date_to_convert, date_to_output)
    If IsDate(date_to_convert) = TRUE Then
        intv_date_mo = DatePart("m", date_to_convert)
        intv_date_day = DatePart("d", date_to_convert)
        intv_date_yr = DatePart("yyyy", date_to_convert)
        date_to_output = intv_date_day & "/" & intv_date_mo & "/" & intv_date_yr
    End If
end function

Function HCRE_panel_bypass()
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function
'===========================================================================================================================

EMConnect ""

forms_to_arep = ""
forms_to_swkr = ""

Call MAXIS_case_number_finder (MAXIS_case_number)
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
today_mo = DatePart("m", date)
today_mo = right("00" & today_mo, 2)

today_day = DatePart("d", date)
today_day = right("00" & today_day, 2)

today_yr = DatePart("yyyy", date)
today_yr = right(today_yr, 2)

today_date = today_mo & "/" & today_day & "/" & today_yr

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 181, 85, "Select the Notice to Send"
  EditBox 60, 5, 50, 15, MAXIS_case_number
  EditBox 45, 25, 130, 15, worker_signature
  'DropListBox 5, 45, 100, 45, "Select One..."+chr(9)+"RECERT - APPT Notice"+chr(9)+"RECERT - NOMI"+chr(9)+"RECERT - VERIFS"+chr(9)+"APPLICATION - APPT Notice"+chr(9)+"APPLICATION - NOMI", memo_to_send
  DropListBox 5, 45, 100, 45, "Select One..."+chr(9)+"RECERT - APPT Notice"+chr(9)+"RECERT - NOMI", memo_to_send
  DropListBox 5, 65, 100, 45, "English"+chr(9)+"Somali"+chr(9)+"Spanish"+chr(9)+"Hmong"+chr(9)+"Russian", select_language
  ButtonGroup ButtonPressed
    OkButton 125, 45, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 35, 10, "Signature"
EndDialog

Do
    Do
        err_msg = ""

        dialog Dialog1
        cancel_without_confirmation
        if IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "Invalid MAXIS Case Number"
        if len(MAXIS_case_number) > 8 Then err_msg = err_msg & vbNewLine & "Invalid MAXIS Case Number"
        If memo_to_send = "Select One..." Then err_msg = err_msg & vbNewLine & "Pick Notice"

        If err_msg <> "" Then MsgBox "Fix:" & vbNewLine & err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


written_lang = "99"         '07, 01, 02, 06, 99'
Select Case select_language
    Case "English"
        written_lang = "99"
    Case "Somali"
        written_lang = "07"
    Case "Spanish"
        written_lang = "01"
    Case "Hmong"
        written_lang = "02"
    Case "Russian"
        written_lang = "06"
End Select

'PROG to determine programs active
Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen cash_prog_one, 2, 6, 67               'reading for active MFIP program - which has different requirements
EMReadScreen cash_stat_one, 4, 6, 74
EMReadScreen cash_prog_two, 2, 7, 67
EMReadScreen cash_stat_two, 4, 7, 74

'MFIP is defaulted to FALSE and will only be changed if PROG reads MFIP as active
If cash_prog_one = "MF" AND cash_stat_one = "ACTV" then MFIP_case = TRUE
If cash_prog_two = "MF" AND cash_stat_two = "ACTV" then MFIP_case = TRUE

EMReadScreen snap_status, 4, 10, 74                'reading the status of SNAP

'SNAP is defaulted to TRUE and will only be changed to FALSE if the status us not active or pending
If snap_status = "ACTV" then SNAP_case = TRUE
If snap_status = "PEND" then SNAP_case = TRUE

Call HCRE_panel_bypass

if MFIP_case = TRUE then           'setting the language for the notices - MFIP or SNAP
    if SNAP_case = TRUE then
        programs = "MFIP/SNAP"
    else
        programs = "MFIP"
    end if
else
    programs = "SNAP"
end if

If memo_to_send = "RECERT - APPT Notice" Then
    'OD Recertifications - APPOINTMENT NOTICE

    month_plus_two = CM_plus_2_mo & "/01/" & CM_plus_2_yr
    last_day_of_recert = DateAdd("d", -1, month_plus_two)
    interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr

    interview_end_date = interview_end_date & ""
    last_day_of_recert = last_day_of_recert & ""

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 191, 105, "Appointment Letter Details"
      EditBox 110, 10, 65, 15, interview_end_date
      EditBox 110, 35, 65, 15, last_day_of_recert
      ButtonGroup ButtonPressed
        OkButton 85, 85, 50, 15
        CancelButton 135, 85, 50, 15
      Text 30, 40, 65, 10, "Last Day of Cert Pd"
      Text 10, 15, 90, 10, "Date of Interview Deadline"
      Text 10, 65, 165, 10, "Sending a APPT Notc for the " & CM_plus_2_mo & "/" & CM_plus_2_yr & " ER in " & select_language
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1
            cancel_without_confirmation
            if IsDate(interview_end_date) = FALSE Then err_msg = err_msg & vbNewLine & "Need a valid date for Interview Deadline"
            if IsDate(last_day_of_recert) = FALSE Then err_msg = err_msg & vbNewLine & "Need a valid date for the last day of Recert"

            If err_msg <> "" Then MsgBox "Fix:" & vbNewLine & err_msg
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    Call navigate_to_MAXIS_screen("SPEC", "MEMO")
    Call start_a_new_spec_memo

    Select Case written_lang


        Case "07"   'Somali (2nd)
            'MsgBox "SOMALI"
            CALL write_variable_in_SPEC_MEMO("Waaxda Adeegyada Aadanaha waxay kuu soo dirtay baakad warqado ah. Waraaqahani waxay cusbooneysiiyaan kiiskaaga " & programs & ".")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Fadlan saxiix, taariikhdana ku qor oo soo celi waraaqaha cusboonaysiinta" & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Waa inaad sidoo kale buuxusaa wareysiga " & programs & "-gaaga si kiisku u sii socdo.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("*Fadlan dhammaystir wareysigaaga inta ka horreysa " & interview_end_date & "*")
            CALL write_variable_in_SPEC_MEMO("Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Kiiskaaga " & programs & " wuxuu xirmi doonaa " & last_day_of_recert & " haddii *")
            CALL write_variable_in_SPEC_MEMO("* aynan helin waraaqahaaga iyo dhamaystirka wareysiga. *")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' CALL write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' CALL write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            Call write_variable_in_SPEC_MEMO(" ")
            CALL write_variable_in_SPEC_MEMO("Qoraallada rabshadaha qoysaska waxaad ka heli kartaa https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. Waxaad kaloo codsan kartaa qoraalkan oo warqad ah.")

        Case "01"   'Spanish (3rd)
            'MsgBox "SPANISH"
            CALL convert_date_to_day_first(interview_end_date, day_first_intv_date)
            CALL convert_date_to_day_first(last_day_of_recert, day_first_last_recert)

            CALL write_variable_in_SPEC_MEMO("El Departamento de Servicios Humanos le envio un paquete con papeles. Son los papeles para renovar su caso " & programs & ".")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Por favor, firmelos, coloque la fecha y envie de regreso los papeles para el 08/" & CM_plus_1_mo & "/" & CM_plus_1_yr & ". Tambien debe realizar una entrevista para que continue su caso " & programs & ".")
            CALL write_variable_in_SPEC_MEMO("***Por favor, complete su entrevista para el " & day_first_intv_date & ".***")
            CALL write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes.")
            CALL write_variable_in_SPEC_MEMO("**Su caso " & programs & " sera cerrado el " & day_first_last_recert & " a menos que recibamos sus papeles y realice la entrevista**")
            CALL write_variable_in_SPEC_MEMO("")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' CALL write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300.")
            ' CALL write_variable_in_SPEC_MEMO("Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' CALL write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)")
			' CALL write_variable_in_SPEC_MEMO("")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            Call write_variable_in_SPEC_MEMO(" ")
            CALL write_variable_in_SPEC_MEMO("Los folletos de violencia domestica estan disponibles en")
            CALL write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            CALL write_variable_in_SPEC_MEMO("Tambien puede solicitar una copia en papel.")
        Case "02"   'Hmong (4th)
            'MsgBox "HMONG"
            CALL write_variable_in_SPEC_MEMO("Lub Koos Haum Department of Human Services tau xa ib pob ntawv tuaj rau koj sent. Cov ntawv no yog tuaj tauj koj txoj kev pab " & programs & ".")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Thov kos npe, tso sij hawm thiab muaj xa cov ntawv tauj rov qab tuaj ua ntej " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Koj yuav tsum mus xam phaj txog koj cov kev pab " & programs & " mas thiaj li tauj tau.")
            CALL write_variable_in_SPEC_MEMO("     *** Thov mus xam phaj ua ntej " & interview_end_date & ". ***")
            CALL write_variable_in_SPEC_MEMO("Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Mon txog Fri.")
            CALL write_variable_in_SPEC_MEMO("**    Koj cov kev pab " & programs & " yuav muab kaw thaum     **")
            CALL write_variable_in_SPEC_MEMO("** " & last_day_of_recert & " tsis li mas peb yuav tsum tau txais koj cov **")
            CALL write_variable_in_SPEC_MEMO("**      ntaub ntawvthiab koj txoj kev xam phaj.         **")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' CALL write_variable_in_SPEC_MEMO("  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' CALL write_variable_in_SPEC_MEMO(" (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
            CALL write_variable_in_SPEC_MEMO("")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            Call write_variable_in_SPEC_MEMO(" ")
            CALL write_variable_in_SPEC_MEMO("Cov ntaub ntawv qhia txog kev raug tsim txom los ntawm cov txheeb ze kuj muaj nyob rau ntawm")
            CALL write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            CALL write_variable_in_SPEC_MEMO("Koj kuj thov tau ib qauv thiab.")

        ' Case "06"   'Russian (5th)
        '     'MsgBox "RUSSIAN"
        '     CALL write_variable_in_SPEC_MEMO("Otdel soczial'ny'x sluzhb otpravil vam paket dokumentaczii.")
        '     CALL write_variable_in_SPEC_MEMO("E'ti dokumenty' dlya obnovleniya vashego " & programs & " dela.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("Podpishite, ukazhite datu i vernite dokumenty' o prodlenii do " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Vy' takzhe dolzhny' projti sobesedovanie dlya prodleniya svoego " & programs & " dela.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("*** Pozhalujsta, projdite sobesedovanie do " & interview_end_date & ". ***")
        '     CALL write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("**    Vash delo " & programs & " zakroetsya " & last_day_of_recert & ", za    **")
        '     CALL write_variable_in_SPEC_MEMO("** isklyucheniem esli my' poluchim vashi dokumenty'  **")
        '     CALL write_variable_in_SPEC_MEMO("**          i vy' projdyote sobesedobanie.           **")
        '     CALL write_variable_in_SPEC_MEMO("   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
        '     Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
        '     Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
        '     Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
        '     Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
        '     Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
        '     Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
        '     CALL write_variable_in_SPEC_MEMO("(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
        '     CALL write_variable_in_SPEC_MEMO("Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")
        ' Case "12"   'Oromo (6th)
        '     'MsgBox "OROMO"
        ' Case "03"   'Vietnamese (7th)
        '     'MsgBox "VIETNAMESE"
        Case Else  'English (1st)
            'MsgBox "ENGLISH"
            CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & programs & " case to continue.")
            CALL write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("  *** Please complete your interview by " & interview_end_date & ". ***")
            Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
            Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("**  Your " & programs & " case will close on " & last_day_of_recert & " unless    **")
            CALL write_variable_in_SPEC_MEMO("** we receive your paperwork and complete the interview. **")
            CALL write_variable_in_SPEC_MEMO("")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
			' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
			' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
			' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
			' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
			' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
			' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
			' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
			' Call write_variable_in_SPEC_MEMO(" ")
			' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            Call write_variable_in_SPEC_MEMO(" ")
            CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")

    End Select

    PF4         'Submit the MEMO

    memo_row = 7                                            'Setting the row for the loop to read MEMOs
    notc_confirm = FALSE         'Defaulting this to 'N'
    Do
        EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
        EMReadScreen print_status, 7, memo_row, 67
        If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
            notc_confirm = TRUE             'If we've found this then no reason to keep looking.
            successful_notices = successful_notices + 1                 'For statistical purposes
            Exit Do
        End If

        memo_row = memo_row + 1           'Looking at next row'
    Loop Until create_date = "        "

    if notc_confirm = TRUE then         'IF the notice was confirmed a CASENOTE will be entered

        Call start_a_new_spec_memo

        Select Case written_lang       'Sending notice by language if possible

        Case "07"   'Somali (2nd)
            'MsgBox "SOMALI"
            CALL write_variable_in_SPEC_MEMO("Nidaamka dib-u-cusboonaysiinta waxaa qayb ka ah inaan heno dhammaan xaqiijinta macaluumaadka. Si loo dedejiyo nidaamka dib-u-cusboonaysiinta, fadlan soo raacicaddaymnaha waraaqaha dib-u-cusboonaysiinta.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Tusaalooyinka caddaynta dakhliga: Qaybta dambe ee")
            CALL write_variable_in_SPEC_MEMO("  jeegaga, qoraalka loo shaqeeyaha, warbixinta dakhliga,")
            CALL write_variable_in_SPEC_MEMO("  xisaabaadka ganacsiga, foomamka canshuurta dakhliga, iwm.")
            CALL write_variable_in_SPEC_MEMO("  * Haddii shaqo kaa dhammaatay, soo dir caddeynta")
            CALL write_variable_in_SPEC_MEMO("    dhamaadka shaqada iyo mushaharka ugu dambeeya.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Tusaalooyinka caddaynta Kharashaadka guryaha (haddii wax")
            CALL write_variable_in_SPEC_MEMO("  isbeddelay): kirada/guriga rasiidka lacag bixinta,")
            CALL write_variable_in_SPEC_MEMO("  bixinta, amaah guri, ijaarka, kabitaanka, iwm.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Tusaalooyinka caddaymaha kharashka caafimaadka (haddii")
            CALL write_variable_in_SPEC_MEMO("  wax isbeddelay): wargadda daawada dhaktarka iyo biilal")
            CALL write_variable_in_SPEC_MEMO("  caafimaad, iwm.")
            CALL write_variable_in_SPEC_MEMO("")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("Haddii aad qabto su'aalo ku saabsan nooca xaqiijinta loo baahan yahay, wac 612-596-1300 qof ayaa ku caawin doona.")

        Case "01"   'Spanish (3rd)
            'MsgBox "SPANISH"
            CALL write_variable_in_SPEC_MEMO("Como parte del Proceso de Renovacion, debemos recibir una verificacion reciente de su informacion. Para acelerar el proceso de renovacion, por favor, envie pruebas de sus papeles de renovacion.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Ejemplos de pruebas de ingresos: resumenes de pagos,")
            CALL write_variable_in_SPEC_MEMO("  declaracion del empleador, reportes de ingresos, libros")
            CALL write_variable_in_SPEC_MEMO("  de contabilidad, formularios de impuestos, etc.")
            CALL write_variable_in_SPEC_MEMO("  * Si un trabajo se ha terminado, envie pruebas de dicha")
            CALL write_variable_in_SPEC_MEMO("    situacion y el ultimo pago.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Ejemplos de pruebas de costos de vivienda (si cambio):")
            CALL write_variable_in_SPEC_MEMO("  recibo de la renta/casa, hipoteca, prestamo, subsidio,")
            CALL write_variable_in_SPEC_MEMO("  etc.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Ejemplos de pruebas de gastos medicos (si cambio):")
            CALL write_variable_in_SPEC_MEMO("  prescripciones y cuentas medicas, etc.")
            CALL write_variable_in_SPEC_MEMO("")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("Si tiene preguntas sobre el tipo de verificacion necesaria, llame al 612-596-1300 y alguien lo/la asistira.")
        Case "02"   'Hmong (4th)
            'MsgBox "HMONG"
            CALL write_variable_in_SPEC_MEMO("Raws li peb txoj kev Rov Tauj Dua mas peb yuav tsum tau txais cov xov tseem ceeb los ntawm koj. Yuav kom tauj tau sai, thov xa cov pov thawj nrog koj ntaub ntawv tauj dua tshiab.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Piv txwv pov thawj txog nyiaj txiag: cov tw tshev, ntawv")
            CALL write_variable_in_SPEC_MEMO("  tom chaw ua hauj lwm, ntawv qhia txog nyiaj txiag, ntawv")
            CALL write_variable_in_SPEC_MEMO("  ua lag luam, ntawv ua se, lwm yam.")
            CALL write_variable_in_SPEC_MEMO("  *Yog hais tias koj txoj hauj lwm tu lawm, xa pav thawj")
            CALL write_variable_in_SPEC_MEMO("   txog hnub kawg ua hauj lwm thiab daim tshev uas yog daim")
            CALL write_variable_in_SPEC_MEMO("   kawg.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Piv txwv cov pov thawj them nqi tsev(yog hais tias")
            CALL write_variable_in_SPEC_MEMO("  hloov): pov thawj xauj tsev/them tsev, ntawv them tuam")
            CALL write_variable_in_SPEC_MEMO("  txhab qiv nyiaj yuav tsev, ntawv cog lus xauj tsev, ntawv")
            CALL write_variable_in_SPEC_MEMO("  them tsev luam, lwm yam.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Piv txwv cov pov thawj txog nqi kho mob(yog hais tias")
            CALL write_variable_in_SPEC_MEMO("  hloov lawm): Ntawv yuav tshuaj thiab nqi kho mob, lwm")
            CALL write_variable_in_SPEC_MEMO("  yam.")
            CALL write_variable_in_SPEC_MEMO("")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("Yog hais tias koj muaj lus nug txog cov yuav tsum muaj cov pov thaqwj twg, hu 612-596-1300 ces neeg mam los pab koj.")

        ' Case "06"   'Russian (5th)
        '     'MsgBox "RUSSIAN"
        '     CALL write_variable_in_SPEC_MEMO("V czelyax obnovleniya proczessa my' dolzhny' poluchit' podtverzhdenie vashej unformaczii.  Chtoby' uskorit' proczess obnovlenie, pozhalujsta, otprav'te dokazatel'stva s vashej dokumentacziej na obnovlenie.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'stv doxoda: koreshki chekov,")
        '     CALL write_variable_in_SPEC_MEMO("  zayavlenie rabotodatelya, otchety' o doxodax,")
        '     CALL write_variable_in_SPEC_MEMO("  buxgalterskie knigi, formy' podoxodnogo naloga i t.d.")
        '     CALL write_variable_in_SPEC_MEMO("  * Esli vy' prekratili rabotat', otprav'te podtberzhdenie")
        '     CALL write_variable_in_SPEC_MEMO("    o prekrashhenii raboty' i poslednyuyu oplatu.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'stv stoimosti zhil'ya (esli oni")
        '     CALL write_variable_in_SPEC_MEMO("  ezmeneny'): arenda/dom kvitancziya ob oplate, ipoteka,")
        '     CALL write_variable_in_SPEC_MEMO("  arenda, subsidiya i t.d.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'ctv mediczinskix rassxodov (esli oni")
        '     CALL write_variable_in_SPEC_MEMO("  izmeneny'): oplata za lekarstva i medeczinskie scheta i")
        '     CALL write_variable_in_SPEC_MEMO("  t. d.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("Esli u vas est' voprosy' o tipe dokazatel'stv pozvonite po telefonu 612-596-1300, u kto-to pomozhet vam.")
        ' Case "12"   'Oromo (6th)
        '     'MsgBox "OROMO"
        ' Case "03"   'Vietnamese (7th)
        '     'MsgBox "VIETNAMESE"
        Case Else  'English (1st)
            'MsgBox "ENGLISH"
            CALL write_variable_in_SPEC_MEMO("As a part of the Renewal Process we must receive recent verification of your information. To speed the renewal process, please send proofs with your renewal paperwork.")
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
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

        End Select

        PF4

        start_a_blank_case_note
        EMSendKey("*** Notice of " & programs & " Recertification Interview Sent ***")
        CALL write_variable_in_case_note("* A notice has been sent to client with detail about how to call in for an interview.")
        CALL write_variable_in_case_note("* Client must submit paperwork and call 612-596-1300 to complete interview.")
        If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
        If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
        call write_variable_in_case_note("---")
        CALL write_variable_in_case_note("Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
        call write_variable_in_case_note("---")
        call write_variable_in_case_note(worker_signature)

        PF3

    ENd If

ElseIf memo_to_send = "RECERT - NOMI" Then
    'OD Recertifications - NOMI

    MAXIS_footer_month = CM_plus_1_mo       'need to look at stat for next month to see if app is received.
    MAXIS_footer_year = CM_plus_1_yr

    Call navigate_to_MAXIS_screen("STAT", "REVW")

    recvd_appl = TRUE

    EmReadscreen caf_recvd_date, 8, 13, 37
    caf_recvd_date = replace(caf_recvd_date, " ", "/")
    If caf_recvd_date = "__/__/__" Then
        recvd_appl = FALSE
        date_of_app = ""
    Else
        date_of_app = caf_recvd_date
    End If

    month_plus_one = CM_plus_1_mo & "/01/" & CM_plus_1_yr
    last_day_of_recert = DateAdd("d", -1, month_plus_one)

    date_of_app = date_of_app & ""
    last_day_of_recert = last_day_of_recert & ""

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 191, 105, "NOMI Details"
      EditBox 110, 10, 65, 15, date_of_app
      EditBox 110, 35, 65, 15, last_day_of_recert
      ButtonGroup ButtonPressed
        OkButton 85, 85, 50, 15
        CancelButton 135, 85, 50, 15
      Text 30, 40, 65, 10, "Last Day of Cert Pd"
      Text 10, 15, 90, 10, "Date Application Received"
      Text 10, 65, 165, 10, "Sending a NOMI for the " & CM_plus_1_mo & "/" & CM_plus_1_yr & " ER in " & select_language
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1
            cancel_without_confirmation
            if IsDate(last_day_of_recert) = FALSE Then err_msg = err_msg & vbNewLine & "Need a valid date for the last day of Recert."
            If IsDate(date_of_app) = FALSE AND trim(date_of_app) <> "" Then err_msg = err_msg & vbNewLine & "Need a valid dateapp received date."
            If err_msg <> "" Then MsgBox "Fix:" & vbNewLine & err_msg
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

    date_of_app = trim(date_of_app)
    If date_of_app <> "" Then recvd_appl = TRUE

    Call navigate_to_MAXIS_screen("SPEC", "MEMO")
    Call start_a_new_spec_memo

    Select Case written_lang       'selecting  the language and writing the memo by language

        Case "07"   'Somali (2nd)
            'MsgBox "SOMALI"
            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("Waraaqahaagii dib-u-cusboonaysiinta waxaan helnay" & date_of_app & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Waraaqahaagii dib-u-cusboonaysiinta weli ma aynaan helin.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Waa inaad wareysi martaa inta ka horreysa " & last_day_of_recert & " haddii kale waxaa joogsan doona waxtarrada aad hesho.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha.")
            CALL write_variable_in_SPEC_MEMO("")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' CALL write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' CALL write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Haddii aynaan war kaa helin inta ka horreysa " & last_day_of_recert & " *")
            CALL write_variable_in_SPEC_MEMO("*   Macaawinada aad hesho waxay instaageysaa " & last_day_of_recert & ".  *")

        Case "01"   'Spanish (3rd)
            'MsgBox "SPANISH"
            CALL convert_date_to_day_first(date_of_app, day_first_app_date)
            CALL convert_date_to_day_first(last_day_of_recert, day_first_last_recert)

            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("Recibimos sus papeles de recertificacion el " & day_first_app_date & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Aun no se han recibido sus Papeles de Recertificacion.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Debe realizar una entrevista para el " & day_first_last_recert & " o sus beneficios se terminaran.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes.")
            CALL write_variable_in_SPEC_MEMO("")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' CALL write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300.")
            ' CALL write_variable_in_SPEC_MEMO("Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' CALL write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("**Si no tenemos novedades suyas para el " & day_first_last_recert & ", sus beneficios se terminaran el " & day_first_last_recert & "**")

        Case "02"   'Hmong (4th)
            'MsgBox "HMONG"
            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("Peb twb txais tau koj cov Ntaub Ntawv Rov Qab Tauj Dua thaum " & date_of_app & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Peb tsis tau txais koj cov Ntaub Ntawv Rov Qab Tauj Duu.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Koj yuav tsum mus xam pphaj ua ntej " & last_day_of_recert & " los yog yuav txiav koj cov kev pab.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Monday txog Friday.")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' CALL write_variable_in_SPEC_MEMO("  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' CALL write_variable_in_SPEC_MEMO(" (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("** Yog hais tias tsis hnov koj teb ua ntej " & last_day_of_recert & "  **")
            CALL write_variable_in_SPEC_MEMO("**   koj cov kev pab yuav raug kaw thaum " & last_day_of_recert & ".   **")

        ' Case "06"   'Russian (5th)
        '     'MsgBox "RUSSIAN"
        '     if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("My' poluchili vashu dokumentacziyu o pereodicheskoj attestaczii " & date_of_app & ".")
        '     if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Vasha dokumentacziya o pereodicheskoj attestaczii eshhyo ne poluchena.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("Vy' dolzhny' projti sobesedovanie do " & last_day_of_recert & " ili vasha programma zakroetsya.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("Chtoby' projti sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
        '     Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
        '     Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
        '     Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
        '     Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
        '     Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
        '     Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
        '     CALL write_variable_in_SPEC_MEMO("(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
        '     CALL write_variable_in_SPEC_MEMO("")
        '     CALL write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vas do " & last_day_of_recert & " **")
        '     CALL write_variable_in_SPEC_MEMO("**   vasha programma zakroetsya " & last_day_of_recert & "    **")

        ' Case "12"   'Oromo (6th)
        '     'MsgBox "OROMO"
        ' Case "03"   'Vietnamese (7th)
        '     'MsgBox "VIETNAMESE"
        Case Else  'English (1st)
            'MsgBox "ENGLISH"
            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("We received your Recertification Paperwork on " & date_of_app & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Your Recertification Paperwork has not yet been received.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
            CALL write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
            Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
            CALL write_variable_in_SPEC_MEMO("")
			'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
            ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
            ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
			CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
            CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")

    End Select

    PF4         'Submit the MEMO

    memo_row = 7                                            'Setting the row for the loop to read MEMOs
    notc_confirm = FALSE         'Defaulting this to 'N'
    Do
        EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
        EMReadScreen print_status, 7, memo_row, 67
        If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
            notc_confirm = TRUE             'If we've found this then no reason to keep looking.
            successful_notices = successful_notices + 1                 'For statistical purposes
            Exit Do
        End If

        memo_row = memo_row + 1           'Looking at next row'
    Loop Until create_date = "        "

    If notc_confirm = TRUE then         'IF the notice was confirmed a CASENOTE will be entered
        start_a_blank_case_note

        EMSendKey("*** NOMI Sent for SNAP Recertification***")
        if recvd_appl = TRUE then CALL write_variable_in_CASE_NOTE("* Recertification app received on " & date_of_app)
        if recvd_appl = FALSE then CALL write_variable_in_CASE_NOTE("* Recertification app has NOT been received. Client must submit paperwork.")
        CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about how to call in for an interview.")
        CALL write_variable_in_CASE_NOTE("* Client must call 612-596-1300 to complete interview.")
        If forms_to_arep = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")
        If forms_to_swkr = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")
        call write_variable_in_case_note("---")
        call write_variable_in_case_note(worker_signature)
        PF3

    End If


ElseIf memo_to_send = "APPLICATION - APPT Notice" Then  'NOT CURRENTLY USED AS THESE HAVE THEIR OWN SCRIPT
    'OD Application - APPOINTMENT NOTICE
    application_date = DateAdd("d", -2, date)
    need_intv_date = DateAdd("d", 5, date)
    last_contact_day = DateAdd("d", 28, date)

    Select Case written_lang
        Case "07"   'Somali (2nd)
            Call write_variable_in_SPEC_MEMO("Taariikhdu markey ahayd " & application_date & ", Waxaad Degmada Hennepin ka codsataycaawimaad, waxaasw loo baahan yahay wareysi si loo hiregeliyo codsigaaga.")
            Call write_variable_in_SPEC_MEMO("** Wareysiga waa in la dhammaystiro ka hor " & need_intv_date & " **")
            Call write_variable_in_SPEC_MEMO("Si loo dhammaystiro wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 aroornimo ilaa 4:30 galabnimo Isniina ilaa Jimcaha.")
            Call write_variable_in_SPEC_MEMO("* Waxaa dhici karta in lagu siiyo gargaarka SNAP 24 saac gudahood wareysiga kaddib.")
            Call write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
            Call write_variable_in_SPEC_MEMO("* Haddii aynaan war kaa helin inta ka horreyssa " & last_contact_day & "*")
            Call write_variable_in_SPEC_MEMO("*              codsigaaga waa la diidi doonaa             *")
            Call write_variable_in_SPEC_MEMO("Haddii aad codsaneyso barnaamijka lacagta caddaanka ah ee haweenka uurka leh ama caruurta yar yar, waxaa laga yaabaa inaad u baahato wareysi fool-ka-fool ah.")
            Call write_variable_in_SPEC_MEMO("Qoraallada rabshadaha qoysaska waxaad ka heli kartaa")
            Call write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("Waxaad kaloo codsan kartaa qoraalkan oo warqad ah.")

            'MsgBox "Somali"

        Case "01"   'Spanish (3rd)

            Call write_variable_in_SPEC_MEMO("Usted ha aplicado para recibir ayuda en el Condado de Hennepin el " & application_date & " y se requiere una entrevista para procesar su aplicacion.")
            Call write_variable_in_SPEC_MEMO("**La entrevista debe ser completada para el " & need_intv_date & ".**")
            Call write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 8:00 a.m. y las 4:30 p.m. de lunes a viernes.")
            Call write_variable_in_SPEC_MEMO("*Puede recibir los beneficios de SNAP dentro de las 24 horas de realizada la entrevista.")
            Call write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300. Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos que se remarque lo contrario)")
            Call write_variable_in_SPEC_MEMO(" **   Si no tenemos novedades suyas para el " & last_contact_day & "   **")
            Call write_variable_in_SPEC_MEMO(" **          su aplicacion sera denegada           **")
            Call write_variable_in_SPEC_MEMO("Si esta aplicando para un programa para mujeres embarazadas o para ninos menores, podria necesitar una entrevista en persona.")
            'Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Los folletos de violencia domestica estan disponibles en")
            Call write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("Tambien puede solicitar una copia en papel.")

            'MsgBox "Spanish"

        Case "02"   'Hmong (4th)
            Call write_variable_in_SPEC_MEMO("Koj tau thov kev pab cuam los ntawm Hennepin County rau thaum " & application_date & " Es yuav tsum tau tuaj xam phaj mas thiaj li yuav khiav koj cov ntaub ntawv.")
            Call write_variable_in_SPEC_MEMO("** Txoj kev xam phaj mas yuav tsum tshwm sim ua ntej lub " & need_intv_date & ". **")
            Call write_variable_in_SPEC_MEMO("Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Mon txog Fri.")
            Call write_variable_in_SPEC_MEMO("* Koj yuav tsim nyob tau cov kev pab SNAP uas siv tau 24 teev tom qab kev xam phaj.")
            Call write_variable_in_SPEC_MEMO(" Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
            Call write_variable_in_SPEC_MEMO("* Yog hais tias peb tsis hnov koj teb ua ntej " & last_contact_day & "*")
            Call write_variable_in_SPEC_MEMO("*           yuav tsis lees koj daim ntawv thov.         *")
            Call write_variable_in_SPEC_MEMO("Yog hais tias koj thov nyiaj ntsuab rau cov poj niam uas cev xeeb tub los yog rau cov menyuam yaus, koj yuav tsum tuaj xam phaj tim ntsej muag.")
            Call write_variable_in_SPEC_MEMO("   Cov ntaub ntawv qhia txog kev raug tsim txom los ntawm cov txheeb ze kuj muaj nyob rau ntawm")
            Call write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("Koj kuj thov tau ib qauv thiab.")

            'MsgBox "Hmong"

        ' Case "06"   'Russian (5th)
        '     Call write_variable_in_SPEC_MEMO("Vy' obratilis' za pomosh'ju v okrug Xennepin " & application_date & " u dlya obrabotki zayavleniya trebuetsya sobesedovanie.")
        '     Call write_variable_in_SPEC_MEMO("** Sobesedovanie dolozhno by't' zaversheno k " & need_intv_date & ". ** ")
        '     Call write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonom, pozbonite v Informaczionnuju liniju EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
        '     Call write_variable_in_SPEC_MEMO("** Vy' smozhete poluchit' vy'platu SNAP vtechenie 24 chasov posle niterv'ju.")
        '     Call write_variable_in_SPEC_MEMO("")
        '     Call write_variable_in_SPEC_MEMO("Esli vy' xotite naznachit' sobesedovanie pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v ljubojiz shesti oficov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
        '     Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
        '     Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
        '     Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
        '     Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
        '     Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
        '     Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
        '     Call write_variable_in_SPEC_MEMO("(Chasy priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
        '     Call write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vac do " & last_contact_day & " **")
        '     Call write_variable_in_SPEC_MEMO("**     vashi zayavlenie budet otklonino.      **")
        '     Call write_variable_in_SPEC_MEMO("Esli vy' podaete zayavku na poluchenie denezhnoj programmy' dlya beremenny'x zhenshhin ili nesovershennoletnix detej, vam mozhet potrebovat'sya lechnoe sobesedobanie.")
        '     Call write_variable_in_SPEC_MEMO("")
        '     Call write_variable_in_SPEC_MEMO("Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG")
        '     Call write_variable_in_SPEC_MEMO("Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")

            'MsgBox "Russian"

        ' Case "12"   'Oromo (6th)
        '     'MsgBox "OROMO"
        ' Case "03"   'Vietnamese (7th)
        '     'MsgBox "VIETNAMESE"
        Case Else  'English (1st)

            Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
            Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & need_intv_date & ". **")
            Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
            Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
            Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **")
            Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")

            'MsgBox "English"
    End Select

ElseIf memo_to_send = "APPLICATION - NOMI" Then         'NOT CURRENTLY USED AS THESE HAVE THEIR OWN SCRIPT'
    'OD Application - NOMI
    application_date = DateAdd("d", -7, date)
    appointment_date = date
    nomi_last_contact_day = DateAdd("d", 23, date)

    Select Case written_lang
        Case "07"   'Somali (2nd)
            Call write_variable_in_SPEC_MEMO("Waxdhawaan dalbatay caawinaad taariikhdu markay ahayd " & application_date & ".")
            Call write_variable_in_SPEC_MEMO("Wareysigaagu wuxuu ahaa in la dhammaystiro ka hor " & appointment_date & ".")
            Call write_variable_in_SPEC_MEMO("Wareysi ayaa loo baahan yahay is loo hirgeliyo codsigaaga.")
            Call write_variable_in_SPEC_MEMO("Si aad u dhamaystirto wareysiga telefoonka, wac laynka taleefanka EZ 612-596-1300 inta u dhaxaysa 8:00 subaxnimo ilaa 4:30 galabnimo Isniinta ilaa Jimcaha.")
            Call write_variable_in_SPEC_MEMO("* Waxaa dhici karta in lagu siiyo gargaarka SNAP 24 saac gudahood wareysiga kaddib.")
            Call write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
            Call write_variable_in_SPEC_MEMO("* Haddii aynaan war kaa helin inta ka horreyssa " & nomi_last_contact_day & " *")
            Call write_variable_in_SPEC_MEMO("*              codsigaaga waa la diidi doonaa             *")
            Call write_variable_in_SPEC_MEMO("Haddii aad codsaneyso barnaamijka lacagta caddaanka ah ee haweenka uurka leh ama caruurta yar yar, waxaa laga yaabaa inaad u baahato wareysi fool-ka-fool ah.")
            Call write_variable_in_SPEC_MEMO("Qoraallada rabshadaha qoysaska waxaad ka heli kartaa")
            Call write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("Waxaad kaloo codsan kartaa qoraalkan oo warqad ah.")

            'MsgBox "Somali"

        Case "01"   'Spanish (3rd)

            Call write_variable_in_SPEC_MEMO("Usted ha aplicado recientemente para recibir ayuda en el Condado de Hennepin el " & application_date & ".")
            Call write_variable_in_SPEC_MEMO("Su entrevista debio haber sido realizada para el " & appointment_date)
            Call write_variable_in_SPEC_MEMO("Se requiere una entrevista para procesar su aplicacion.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 8:00 a.m. y las 4:30 p.m. de lunes a viernes.")
            Call write_variable_in_SPEC_MEMO("*Puede recibir los beneficios de SNAP dentro de las 24 horas de realizada la entrevista.")
            Call write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300. Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos que se remarque lo contrario)")
            Call write_variable_in_SPEC_MEMO(" **   Si no tenemos novedades suyas para el " & nomi_last_contact_day & "   **")
            Call write_variable_in_SPEC_MEMO(" **             su aplicacion sera denegada              **")
            Call write_variable_in_SPEC_MEMO("Si esta aplicando para un programa para mujeres embarazadas o para ninos menores, podria necesitar una entrevista en persona.")
            'Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Los folletos de violencia domestica estan disponibles en")
            Call write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("Tambien puede solicitar una copia en papel.")

            'MsgBox "Spanish"

        Case "02"   'Hmong (4th)
            Call write_variable_in_SPEC_MEMO("Koj nyuam qhuav ua ntawv tuaj thov kev pav thaum lub " & application_date & ". Kev yuav xam phaj koj mas yuav tsum tiav hlo rau thaum lub " & appointment_date & ". Yuav tsum muaj kev xam phaj mas thiaj li yuav pib khiav tau koj cov ntaub ntawv.")
            Call write_variable_in_SPEC_MEMO("  Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Mon txog Fri.")
            Call write_variable_in_SPEC_MEMO("* Koj yuav tsim nyob tau cov kev pab SNAP uas siv tau 24 teev tom qab kev xam phaj.")
            Call write_variable_in_SPEC_MEMO(" Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
            Call write_variable_in_SPEC_MEMO("** Yog hais tias peb tsis hnov koj teb ua ntej " & nomi_last_contact_day & "**")
            Call write_variable_in_SPEC_MEMO("**         yuav tsis lees koj daim ntawv thov.          **")
            Call write_variable_in_SPEC_MEMO("Yog hais tias koj thov nyiaj ntsuab rau cov poj niam uas cev xeeb tub los yog rau cov menyuam yaus, koj yuav tsum tuaj xam phaj tim ntsej muag.")
            Call write_variable_in_SPEC_MEMO("   Cov ntaub ntawv qhia txog kev raug tsim txom los ntawm cov txheeb ze kuj muaj nyob rau ntawm")
            Call write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            Call write_variable_in_SPEC_MEMO("Koj kuj thov tau ib qauv thiab.")


        ' Case "06"   'Russian (5th)
        '     Call write_variable_in_SPEC_MEMO("Vy' podali zayavlenie na pomoshh' " & application_date & ".")
        '     Call write_variable_in_SPEC_MEMO("Vashe sobesedovanie dolzhno by't' zaversheno k " & appointment_date & ".")
        '     Call write_variable_in_SPEC_MEMO("Dlya obrabotki zayavleniya trebuetsya sobesedovanie.")
        '     Call write_variable_in_SPEC_MEMO("")
        '     Call write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonom, pozbonite v Informaczionnuju liniju EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
        '     Call write_variable_in_SPEC_MEMO("** Vy' smozhete poluchit' vy'platu SNAP vtechenie 24 chasov posle niterv'ju.")
        '     Call write_variable_in_SPEC_MEMO("")
        '     Call write_variable_in_SPEC_MEMO("Esli vy' xotite naznachit' sobesedovanie pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v ljubojiz shesti oficov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
        '     Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
        '     Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
        '     Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
        '     Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
        '     Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
        '     Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
        '     Call write_variable_in_SPEC_MEMO("(Chasy priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
        '     Call write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vac do " & nomi_last_contact_day & " **")
        '     Call write_variable_in_SPEC_MEMO("**     vashi zayavlenie budet otklonino.     **")
        '     Call write_variable_in_SPEC_MEMO("Esli vy' podaete zayavku na poluchenie denezhnoj programmy' dlya beremenny'x zhenshhin ili nesovershennoletnix detej, vam mozhet potrebovat'sya lechnoe sobesedobanie.")
        '     Call write_variable_in_SPEC_MEMO("  Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG")
        '     Call write_variable_in_SPEC_MEMO("Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")

            'MsgBox "Russian"

        ' Case "12"   'Oromo (6th)
        '     'MsgBox "OROMO"
        ' Case "03"   'Vietnamese (7th)
        '     'MsgBox "VIETNAMESE"
        Case Else  'English (1st)

            Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
            Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & appointment_date & ".")
            Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
            Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
            Call write_variable_in_SPEC_MEMO(" ")
            Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
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

    End Select
End If

script_end_procedure("")
