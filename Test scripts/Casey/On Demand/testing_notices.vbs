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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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

Dialog1 = ""
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

written_lang = "99"         '07, 01, 02, 06, 99'

Call start_a_new_spec_memo

If memo_to_send = "RECERT - APPT Notice" Then
    'OD Recertifications - APPOINTMENT NOTICE

    programs = "MFIP/SNAP"
    last_day_of_recert = CM_plus_1_mo & "/30/" & CM_plus_1_yr
    interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr
    'NOTICE ON LINE 768'
    ' 'EMSendKey("************************************************************")           'for some reason this is more stable then using write_variable
    ' CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & programs & " case to continue.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' Call write_variable_in_SPEC_MEMO("  *** Please complete your interview by " & interview_end_date & ". ***")
    ' Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    ' Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("**  Your " & programs & " case will close on " & last_day_of_recert & " unless    **")
    ' CALL write_variable_in_SPEC_MEMO("** we receive your paperwork and complete the interview. **")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")
    '
    '
    '
    '
    ' 'Call write_variable_in_SPEC_MEMO("Some cases are eligible to have SNAP benefits issued within 24 hours of the interview, call right away if you have an urgent need.")
    ' 'Call write_variable_in_SPEC_MEMO("Interviews can also be completed in person at one of our six offices:")
    '
    ' ' Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
    ' ' Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    ' '
    ' '
    ' '
    ' ' CALL write_variable_in_SPEC_MEMO("We must have your renewal paperwork to do your interview. Please send proofs with your renewal paperwork.")
    ' ' CALL write_variable_in_SPEC_MEMO("")
    ' ' CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, income reports, business ledgers, income tax forms, etc.")
    ' ' CALL write_variable_in_SPEC_MEMO("")
    ' ' CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house payment receipt, mortgage, lease, subsidy, etc.")
    ' ' CALL write_variable_in_SPEC_MEMO("")
    ' ' CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed): prescription and medical bills, etc.")
    ' ' CALL write_variable_in_SPEC_MEMO("")

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
            CALL write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
            CALL write_variable_in_SPEC_MEMO("Qoraallada rabshadaha qoysaska waxaad ka heli kartaa https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. Waxaad kaloo codsan kartaa qoraalkan oo warqad ah.")

        Case "01"   'Spanish (3rd)
            'MsgBox "SPANISH"

            CALL write_variable_in_SPEC_MEMO("El Departamento de Servicios Humanos le envio un paquete con papeles. Son los papeles para renovar su caso " & programs & ".")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Por favor, firmelos, coloque la fecha y envie de regreso los papeles para el 08/" & CM_plus_1_mo & "/" & CM_plus_1_yr & ". Tambien debe realizar una entrevista para que continue su caso " & programs & ".")
            CALL write_variable_in_SPEC_MEMO("***Por favor, complete su entrevista para el " & interview_end_date & ".***")
            CALL write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes.")
            CALL write_variable_in_SPEC_MEMO("**Su caso " & programs & " sera cerrado el " & last_day_of_recert & " a menos que recibamos sus papeles y realice la entrevista**")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300.")
            CALL write_variable_in_SPEC_MEMO("Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)")
            CALL write_variable_in_SPEC_MEMO("")
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
            CALL write_variable_in_SPEC_MEMO("  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO(" (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Cov ntaub ntawv qhia txog kev raug tsim txom los ntawm cov txheeb ze kuj muaj nyob rau ntawm")
            CALL write_variable_in_SPEC_MEMO("https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
            CALL write_variable_in_SPEC_MEMO("Koj kuj thov tau ib qauv thiab.")

        Case "06"   'Russian (5th)
            'MsgBox "RUSSIAN"
            CALL write_variable_in_SPEC_MEMO("Otdel soczial'ny'x sluzhb otpravil vam paket dokumentaczii.")
            CALL write_variable_in_SPEC_MEMO("E'ti dokumenty' dlya obnovleniya vashego " & programs & " dela.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Podpishite, ukazhite datu i vernite dokumenty' o prodlenii do " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". Vy' takzhe dolzhny' projti sobesedovanie dlya prodleniya svoego " & programs & " dela.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("*** Pozhalujsta, projdite sobesedovanie do " & interview_end_date & ". ***")
            CALL write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("**    Vash delo " & programs & " zakroetsya " & last_day_of_recert & ", za    **")
            CALL write_variable_in_SPEC_MEMO("** isklyucheniem esli my' poluchim vashi dokumenty'  **")
            CALL write_variable_in_SPEC_MEMO("**          i vy' projdyote sobesedobanie.           **")
            CALL write_variable_in_SPEC_MEMO("   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO("(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
            CALL write_variable_in_SPEC_MEMO("Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")
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

    End Select

ElseIf memo_to_send = "RECERT - VERIFS" Then

    ' CALL write_variable_in_SPEC_MEMO("As a part of the Renewal Process we must receive recent verification of your information. To speed the renewal process, please send proofs with your renewal paperwork.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, employer statement,")
    ' CALL write_variable_in_SPEC_MEMO("   income reports, business ledgers, income tax forms, etc.")
    ' CALL write_variable_in_SPEC_MEMO("   *If a job has ended, send proof of the end of employment")
    ' CALL write_variable_in_SPEC_MEMO("   and last pay.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house")
    ' CALL write_variable_in_SPEC_MEMO("   payment receipt, mortgage, lease, subsidy, etc.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed):")
    ' CALL write_variable_in_SPEC_MEMO("   prescription and medical bills, etc.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

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
        CALL write_variable_in_SPEC_MEMO("Yog hais tias koj muaj lus nug txog cov yuav tsum muaj cov pov thaqwj twg, hu 612-596-1300 ces neeg mam los pab koj.")

    Case "06"   'Russian (5th)
        'MsgBox "RUSSIAN"
        CALL write_variable_in_SPEC_MEMO("V czelyax obnovleniya proczessa my' dolzhny' poluchit' podtverzhdenie vashej unformaczii.  Chtoby' uskorit' proczess obnovlenie, pozhalujsta, otprav'te dokazatel'stva s vashej dokumentacziej na obnovlenie.")
        CALL write_variable_in_SPEC_MEMO("")
        CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'stv doxoda: koreshki chekov,")
        CALL write_variable_in_SPEC_MEMO("  zayavlenie rabotodatelya, otchety' o doxodax,")
        CALL write_variable_in_SPEC_MEMO("  buxgalterskie knigi, formy' podoxodnogo naloga i t.d.")
        CALL write_variable_in_SPEC_MEMO("  * Esli vy' prekratili rabotat', otprav'te podtberzhdenie")
        CALL write_variable_in_SPEC_MEMO("    o prekrashhenii raboty' i poslednyuyu oplatu.")
        CALL write_variable_in_SPEC_MEMO("")
        CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'stv stoimosti zhil'ya (esli oni")
        CALL write_variable_in_SPEC_MEMO("  ezmeneny'): arenda/dom kvitancziya ob oplate, ipoteka,")
        CALL write_variable_in_SPEC_MEMO("  arenda, subsidiya i t.d.")
        CALL write_variable_in_SPEC_MEMO("")
        CALL write_variable_in_SPEC_MEMO("* Primery' dokazatel'ctv mediczinskix rassxodov (esli oni")
        CALL write_variable_in_SPEC_MEMO("  izmeneny'): oplata za lekarstva i medeczinskie scheta i")
        CALL write_variable_in_SPEC_MEMO("  t. d.")
        CALL write_variable_in_SPEC_MEMO("")
        CALL write_variable_in_SPEC_MEMO("Esli u vas est' voprosy' o tipe dokazatel'stv pozvonite po telefonu 612-596-1300, u kto-to pomozhet vam.")
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
        CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

    End Select

ElseIf memo_to_send = "RECERT - NOMI" Then
    'OD Recertifications - NOMI
    recvd_appl = TRUE
    date_of_app = DateAdd("d", -5, date)
    last_day_of_recert = CM_mo & "/30/" & CM_yr

    'NOTICE ON LINE 902'
    ' if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("We received your Recertification Paperwork on " & date_of_app & ".")
    ' if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Your Recertification Paperwork has not yet been received.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
    ' CALL write_variable_in_SPEC_MEMO("")
    '
    '
    ' Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    ' Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    '
    '
    ' ' CALL write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at 612-596-1300 between 8:00am and 4:30pm Monday through Friday.")
    ' ' CALL write_variable_in_SPEC_MEMO("")
    ' ' CALL write_variable_in_SPEC_MEMO("You may also come in to the office to complete an interview between 8:00 am and 4:30pm Monday through Friday.")
    ' CALL write_variable_in_SPEC_MEMO("")
    ' CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
    ' CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")

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
            CALL write_variable_in_SPEC_MEMO("Haddii aad rabto inaad samaysato ballan wareysi, wac 612-596-1300. Waxa kale oo aad iman kartaa mid ka mid ah lixda xafiis ee hoos ku qoran si loo sameeyo wareysi gof ahaaneed inta u dhexeeya 8 ilaa 4:30, Isniinta ilaa jmcaha.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO("(Saacaduhu waa Isniinta - Jimcaha 8-4:30 haddii aan si kale loo sheegin.)")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("* Haddii aynaan war kaa helin inta ka horreysa " & last_day_of_recert & " *")
            CALL write_variable_in_SPEC_MEMO("*   Macaawinada aad hesho waxay instaageysaa " & last_day_of_recert & ".  *")

        Case "01"   'Spanish (3rd)
            'MsgBox "SPANISH"

            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("Recibimos sus papeles de recertificacion el " & date_of_app & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Aun no se han recibido sus Papeles de Recertificacion.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Debe realizar una entrevista para el " & last_day_of_recert & " o sus beneficios se terminaran.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Para completar una entrevista telefonica, llame a la linea de informacion EZ al 612-596-1300 entre las 9 am y las 4 pm de lunes a viernes.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Si desea programar una entrevista, llame al 612-596-1300.")
            CALL write_variable_in_SPEC_MEMO("Tambien puede acercarse a cualquiera de las seis oficinas mencionadas debajo para tener una entrevista personal entre las 8 y las 4:30 de lunes a viernes.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 J h.: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO("(Los horarios son de lunes a viernes de 8 a 4:30 a menos   que se remarque lo contrario)")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("**Si no tenemos novedades suyas para el " & last_day_of_recert & ", sus beneficios se terminaran el " & last_day_of_recert & "**")

        Case "02"   'Hmong (4th)
            'MsgBox "HMONG"
            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("Peb twb txais tau koj cov Ntaub Ntawv Rov Qab Tauj Dua thaum " & date_of_app & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Peb tsis tau txais koj cov Ntaub Ntawv Rov Qab Tauj Duu.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Koj yuav tsum mus xam pphaj ua ntej " & last_day_of_recert & " los yog yuav txiav koj cov kev pab.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Yog xam phaj hauv xov tooj, hu rau EZ Info Line ntawm 612-596-1300 thaum 8:00am thib 4:30pm hnub Monday txog Friday.")
            CALL write_variable_in_SPEC_MEMO("  Yog hais tias koj xav teem tuaj xam phaj, hu 612-596-1300 Koj kuj tuaj tau rau ib lub ntawm rau lub hoob kas nyob hauv qab no tuaj xam phaj tim ntej muag thaum 8 thiab 4:30, hnub Monday txog Friday.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO(" (Cov sij hawm qhib yog M - F 8-4:30 tsis li mas yuav tsum qhia ua ntej)")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("** Yog hais tias tsis hnov koj teb ua ntej " & last_day_of_recert & "  **")
            CALL write_variable_in_SPEC_MEMO("**   koj cov kev pab yuav raug kaw thaum " & last_day_of_recert & ".   **")

        Case "06"   'Russian (5th)
            'MsgBox "RUSSIAN"
            if recvd_appl = TRUE then CALL write_variable_in_SPEC_MEMO("My' poluchili vashu dokumentacziyu o pereodicheskoj attestaczii " & date_of_app & ".")
            if recvd_appl = FALSE then CALL write_variable_in_SPEC_MEMO("Vasha dokumentacziya o pereodicheskoj attestaczii eshhyo ne poluchena.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Vy' dolzhny' projti sobesedovanie do " & last_day_of_recert & " ili vasha programma zakroetsya.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("Chtoby' projti sobesedovanie po telefonu, pozvonite v Informaczionnuyu liniyu EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("   Esli vy' xotite naznachit' sobesedovanie, pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v lyuboj iz shesti ofisov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            CALL write_variable_in_SPEC_MEMO("(Chasy' priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vas do " & last_day_of_recert & " **")
            CALL write_variable_in_SPEC_MEMO("**   vasha programma zakroetsya " & last_day_of_recert & "    **")

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
            Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
            CALL write_variable_in_SPEC_MEMO("")
            CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
            CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")

    End Select


ElseIf memo_to_send = "APPLICATION - APPT Notice" Then  'THIS ONE IS DONE AND ILSE IS VETTING'
    'OD Application - APPOINTMENT NOTICE
    application_date = DateAdd("d", -2, date)
    need_intv_date = DateAdd("d", 5, date)
    last_contact_day = DateAdd("d", 28, date)

    'NOTICE ON LINE 1113'
    ' 'EMsendkey("************************************************************")
    ' 'Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & application_date & ".")
    ' Call write_variable_in_SPEC_MEMO("You applied for assistance in Hennepin County on " & application_date & "")
    ' Call write_variable_in_SPEC_MEMO("and an interview is required to process your application.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("** The interview must be completed by " & need_intv_date & ". **")
    ' Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    ' 'Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday through Friday.")
    ' Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
    ' 'Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Mon through Fri.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' 'Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday through Friday.")
    ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
    '
    ' 'Call write_variable_in_SPEC_MEMO("Some cases are eligible to have SNAP benefits issued within 24 hours of the interview, call right away if you have an urgent need.")
    ' 'Call write_variable_in_SPEC_MEMO("Interviews can also be completed in person at one of our six offices:")
    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_contact_day & " **")
    ' Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    ' Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    ' Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")
    ' 'Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3).")
    ' 'Call write_variable_in_SPEC_MEMO("************************************************************")

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

        Case "06"   'Russian (5th)
            Call write_variable_in_SPEC_MEMO("Vy' obratilis' za pomosh'ju v okrug Xennepin " & application_date & " u dlya obrabotki zayavleniya trebuetsya sobesedovanie.")
            Call write_variable_in_SPEC_MEMO("** Sobesedovanie dolozhno by't' zaversheno k " & need_intv_date & ". ** ")
            Call write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonom, pozbonite v Informaczionnuju liniju EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
            Call write_variable_in_SPEC_MEMO("** Vy' smozhete poluchit' vy'platu SNAP vtechenie 24 chasov posle niterv'ju.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Esli vy' xotite naznachit' sobesedovanie pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v ljubojiz shesti oficov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Chasy priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
            Call write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vac do " & last_contact_day & " **")
            Call write_variable_in_SPEC_MEMO("**     vashi zayavlenie budet otklonino.      **")
            Call write_variable_in_SPEC_MEMO("Esli vy' podaete zayavku na poluchenie denezhnoj programmy' dlya beremenny'x zhenshhin ili nesovershennoletnix detej, vam mozhet potrebovat'sya lechnoe sobesedobanie.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG")
            Call write_variable_in_SPEC_MEMO("Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")

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

ElseIf memo_to_send = "APPLICATION - NOMI" Then
    'OD Application - NOMI
    application_date = DateAdd("d", -7, date)
    appointment_date = date
    nomi_last_contact_day = DateAdd("d", 23, date)



    ' 'NOTICE ON LINE 1223'
    ' 'EMsendkey("************************************************************")
    ' Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_date & ".")
    ' Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & appointment_date & ".")
    ' Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
    ' Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("* You may be able to have SNAP benefits issued within 24 hours of the interview.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
    ' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
    ' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
    ' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
    ' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
    ' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
    ' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
    ' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & nomi_last_contact_day & " **")
    ' Call write_variable_in_SPEC_MEMO("  **    your application will be denied.     **") 'add 30 days
    ' Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
    ' Call write_variable_in_SPEC_MEMO(" ")
    ' Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    ' Call write_variable_in_SPEC_MEMO("You can also request a paper copy.  Auth: 7CFR 273.2(e)(3).")
    '
    '
    ' '
    ' ' Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at ")
    ' ' Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday through Friday.")
    ' ' Call write_variable_in_SPEC_MEMO(" ")
    ' ' Call write_variable_in_SPEC_MEMO("If you do not complete the interview by " & nomi_last_contact_day & " your application will be denied.") 'add 30 days
    ' ' Call write_variable_in_SPEC_MEMO(" ")
    ' ' Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to- face interview.")
    ' ' Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
    ' ' Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
    ' ' Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
    ' ' Call write_variable_in_SPEC_MEMO("************************************************************")

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


        Case "06"   'Russian (5th)
            Call write_variable_in_SPEC_MEMO("Vy' podali zayavlenie na pomoshh' " & application_date & ".")
            Call write_variable_in_SPEC_MEMO("Vashe sobesedovanie dolzhno by't' zaversheno k " & appointment_date & ".")
            Call write_variable_in_SPEC_MEMO("Dlya obrabotki zayavleniya trebuetsya sobesedovanie.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Chtoby' zavershit' sobesedovanie po telefonom, pozbonite v Informaczionnuju liniju EZ po telefonu 612-596-1300 s 8:00 do 16:30 s ponedel'nika po pyatniczu.")
            Call write_variable_in_SPEC_MEMO("** Vy' smozhete poluchit' vy'platu SNAP vtechenie 24 chasov posle niterv'ju.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Esli vy' xotite naznachit' sobesedovanie pozvonite po telefonu 612-596-1300. Vy' takzhe mozhete obratit'sya v ljubojiz shesti oficov. Dlya sobesedovanie s 8 i do 4:30, s ponedel'nika po pyatniczu.")
            Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
            Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
            Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
            Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
            Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
            Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
            Call write_variable_in_SPEC_MEMO("(Chasy priyoma s ponedel'nika po pyatniczu s 8 do 4:30, esli ne ukazano inoe.)")
            Call write_variable_in_SPEC_MEMO("** Esli my' ne usly'shim ot vac do " & nomi_last_contact_day & " **")
            Call write_variable_in_SPEC_MEMO("**     vashi zayavlenie budet otklonino.     **")
            Call write_variable_in_SPEC_MEMO("Esli vy' podaete zayavku na poluchenie denezhnoj programmy' dlya beremenny'x zhenshhin ili nesovershennoletnix detej, vam mozhet potrebovat'sya lechnoe sobesedobanie.")
            Call write_variable_in_SPEC_MEMO("  Broshyupy' o nasilii v sem'e dostupny' po adresu https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG")
            Call write_variable_in_SPEC_MEMO("Vy' takzhe mozhete zaprosit' bumazhnuyu kopiyu.")

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
