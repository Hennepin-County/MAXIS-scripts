'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - RETURNED MAIL RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
call changelog_update("06/06/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog case_number_dlg, 0, 0, 251, 85, "Returned Mail"
  EditBox 55, 5, 40, 15, maxis_case_number
  EditBox 205, 5, 40, 15, date_received
  DropListBox 55, 25, 190, 15, "Select One:"+chr(9)+"Mail has been returned NO forwarding address"+chr(9)+"Mail has been returned with forwarding address in MN"+chr(9)+"Mail has been returned with forwarding address outside MN"+chr(9)+"Client has not responded to request for SVF", ADDR_actions
  EditBox 130, 45, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 145, 65, 50, 15
    CancelButton 195, 65, 50, 15
  Text 5, 30, 45, 10, "Select action:"
  Text 65, 50, 60, 10, "Worker signature:"
  Text 5, 10, 45, 10, "Case number:"
  Text 150, 10, 50, 10, "Date received:"
EndDialog


DO
    DO
    	err_msg = ""
    	DIALOG case_number_dlg
    		IF ButtonPressed = 0 THEN stopscript
    		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
			If isdate(date_received) = FALSE then err_msg = err_msg & vbnewline & "* Please enter a date (--/--/--) in the footer month that you are working in."
			IF ADDR_actions = "Select One:" THEN err_msg = err_msg & vbCr & "Please chose an action for the returned mail."
    		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
CALL check_for_MAXIS(False)

'MEMB_number = "01"
'actual_date = date & ""
'Determines which programs are currently status_checking in the month of application
CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
'Checking for PRIV cases.
'EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
'IF priv_check = "PRIV" THEN
'	priv_case_list = priv_case_list & "|" & MAXIS_case_number
'ELSE						'For all of the cases that aren't privileged...
'Setting some variables for the loop
CASH_STATUS = FALSE 'overall variable'
CCA_STATUS = FALSE
DW_STATUS = FALSE 'Diversionary Work Program'
ER_STATUS = FALSE
FS_STATUS = FALSE
GA_STATUS = FALSE 'General Assistance'
GRH_STATUS = FALSE
HC_STATUS = FALSE
MS_STATUS = FALSE 'Mn Suppl Aid '
MF_STATUS = FALSE 'Mn Family Invest Program '
RC_STATUS = FALSE 'Refugee Cash Assistance'

'Reading the status and program
EMReadScreen cash1_status_check, 4, 6, 74
EMReadScreen cash2_status_check, 4, 7, 74
EMReadScreen emer_status_check, 4, 8, 74
EMReadScreen grh_status_check, 4, 9, 74
EMReadScreen fs_status_check, 4, 10, 74
EMReadScreen ive_status_check, 4, 11, 74
EMReadScreen hc_status_check, 4, 12, 74
EMReadScreen cca_status_check, 4, 14, 74
EMReadScreen cash1_prog_check, 2, 6, 67
EMReadScreen cash2_prog_check, 2, 7, 67
EMReadScreen emer_prog_check, 2, 8, 67
EMReadScreen grh_prog_check, 2, 9, 67
EMReadScreen fs_prog_check, 2, 10, 67
EMReadScreen ive_prog_check, 2, 11, 67
EMReadScreen hc_prog_check, 2, 12, 67
EMReadScreen cca_prog_check, 2, 14, 67

IF FS_status_check = "ACTV" or FS_status_check = "PEND"  THEN FS_STATUS = TRUE
IF emer_status_check = "ACTV" or emer_status_check = "PEND"  THEN ER_STATUS = TRUE
IF grh_status_check = "ACTV" or grh_status_check = "PEND"  THEN GRH_STATUS = TRUE
IF hc_status_check = "ACTV" or hc_status_check = "PEND"  THEN HC_STATUS = TRUE
IF cca_status_check = "ACTV" or cca_status_check = "PEND"  THEN CCA_STATUS = TRUE
'Logic to determine if MFIP is active
If cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "RC" or cash1_prog_check = "MS" THEN
	If cash1_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash1_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash1_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash1_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash1_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash1_status_check = ""     Then CASH_STATUS = FALSE
END IF
If cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "RC" or cash2_prog_check = "MS" THEN
	If cash2_status_check = "ACTV" Then CASH_STATUS = TRUE
	If cash2_status_check = "PEND" Then CASH_STATUS = TRUE
	If cash2_status_check = "INAC" Then CASH_STATUS = FALSE
	If cash2_status_check = "SUSP" Then CASH_STATUS = FALSE
	If cash2_status_check = "DENY" Then CASH_STATUS = FALSE
	If cash2_status_check = ""     Then CASH_STATUS = FALSE
END IF

IF ADDR_actions = "Mail has been returned NO forwarding address" THEN
    BeginDialog RETURNED_MAIL, 0, 0, 181, 185, "Mail has been returned NO forwarding address"
    CheckBox 10, 70, 70, 10, "Sent DHS-2919A", verifA_sent_checkbox
    CheckBox 85, 70, 65, 10, "Sent DHS-2952", SHEL_form_sent_checkbox
    CheckBox 10, 85, 65, 10, "Sent DHS-2402", CRF_sent_checkbox
    EditBox 110, 105, 65, 15, METS_case_number
    DropListBox 110, 125, 65, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", MNsure_ADDR
    EditBox 50, 145, 125, 15, other_notes
    ButtonGroup ButtonPressed
    OkButton 70, 165, 50, 15
    CancelButton 125, 165, 50, 15
    Text 5, 130, 95, 10, "METS correspondence sent:"
    Text 5, 150, 40, 10, "Other notes:"
    GroupBox 5, 5, 170, 50, "NOTE:"
    GroupBox 5, 60, 170, 40, "Verification Request Form"
    Text 5, 110, 70, 10, "METS case number:"
    Text 10, 15, 160, 35, "Do not make any changes to STAT/ADDR.  Do NOT enter a ? or unknown or other county codes on the ADDR panel.  The ADDR panel is used to mail notices; the post office requires an address. "
    EndDialog

	DO
	    DO
	    	err_msg = ""
	    	DIALOG case_number_dlg
	    		IF ButtonPressed = 0 THEN stopscript
	    		IF MNsure_ADDR = "YES" THEN
					IF METS_case_number = "" OR (METS_case_number <> "" AND len(METS_case_number) > 10) OR (METS_case_number <> "" AND IsNumeric(METS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
				END IF
				IF MNsure_ADDR = "Select One:" THEN err_msg = err_msg & vbCr & "Please select if correspondence has been sent to METS."

	    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	    	LOOP UNTIL err_msg = ""
	    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	CALL check_for_MAXIS(False)
END IF

IF ADDR_actions = "Mail has been returned with forwarding address in MN" THEN
    CALL navigate_to_MAXIS_screen("STAT", "ADDR")
    'Writes spreadsheet info to ADDR
    EMreadscreen ADDR_line_one, 20, 6, 43
    EMreadscreen ADDR_line_two, 20, 7, 43
    EMreadscreen ADDR_city, 15, 8, 43
    'EMreadscreen "MN", 8, 66		'Defaults to MN for all cases at this time
    EMreadscreen ADDR_zip, 5, 9, 43
    EMreadscreen ADDR_county, 2, 9, 66
    EMreadscreen ADDR_addr_verif, 2, 9, 74
    EMreadscreen ADDR_homeless, 1, 10, 43
    'EMreadscreen ADDR_reservation, 1, 10, 74
    EMreadscreen ADDR_mailing_addr_line_one, 20, 13, 43
    EMreadscreen ADDR_mailing_addr_line_two, 20, 14, 43
    EMreadscreen ADDR_mailing_addr_city, 15, 15, 43
    EMreadscreen 2, 16, 43	'Only writes if the user indicated a mailing address. Defaults to MN at this time.
    EMreadscreen ADDR_mailing_addr_zip, 5, 16, 52
    EMreadscreen ADDR_phone_1A, 3, 17, 45						'Has to split phone numbers up into three parts each
    EMreadscreen ADDR_phone_2B, 3, 17, 51
    EMreadscreen ADDR_phone_3C, 4, 17, 55
    EMreadscreen ADDR_phone_2A, 3, 18, 45
    EMreadscreen ADDR_phone_2B, 3, 18, 51
    EMreadscreen ADDR_phone_2C, 4, 18, 55
    EMreadscreen ADDR_phone_3A, 3, 19, 45
    EMreadscreen ADDR_phone_3B, 3, 19, 51
    EMreadscreen ADDR_phone_3C, 4, 19, 55


IF ADDR_actions = "Mail has been returned with forwarding address outside MN" THEN
IF ADDR_actions = "Client has not responded to request for SVF" THEN
BeginDialog RETURNED_MAIL, 0, 0, 185, 335, "RETURNED MAIL DIALOG"
  EditBox 110, 5, 65, 15, date_received
  EditBox 55, 25, 120, 15, from_ADDR
  EditBox 55, 45, 120, 15, from_CITY
  EditBox 55, 65, 20, 15, from_STATE
  EditBox 120, 65, 55, 15, from_ZIP
  DropListBox 115, 85, 35, 15, "No"+chr(9)+"Yes", forwarding_ADDR
  EditBox 55, 105, 120, 15, new_ADDR
  EditBox 55, 125, 120, 15, new_CITY
  EditBox 55, 145, 25, 15, new_STATE
  EditBox 135, 145, 40, 15, new_ZIP
  DropListBox 110, 165, 35, 15, "No"+chr(9)+"Yes", updated_ADDR
  EditBox 110, 180, 65, 15, new_COUNTY
  CheckBox 50, 200, 70, 10, "Sent DHS-2919A", verifA_sent_checkbox
  CheckBox 50, 210, 65, 10, "Sent DHS-2952", SHEL_form_sent_checkbox
  CheckBox 50, 220, 65, 10, "Sent DHS-2402", CRF_sent_checkbox
  DropListBox 120, 230, 30, 15, "No"+chr(9)+"Yes", returned_mail_resent_list
  DropListBox 105, 245, 45, 15, "Select"+chr(9)+"Yes"+chr(9)+"No", MNsure_active
  EditBox 100, 260, 75, 15, MNsure_number
  DropListBox 100, 278, 40, 10, "N/A"+chr(9)+"Yes"+chr(9)+"No", MNsure_ADDR
  EditBox 100, 295, 75, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 5, 315, 50, 15
    CancelButton 130, 315, 50, 15
  Text 45, 235, 70, 10, "Returned mail resent:"
  Text 85, 150, 50, 10, "New Zip Code:"
  Text 10, 300, 90, 10, "Misc notes/Actions Taken:"
  Text 10, 30, 40, 10, "From ADDR:"
  Text 10, 110, 40, 10, "New ADDR:"
  Text 30, 70, 20, 10, "State:"
  Text 10, 150, 40, 10, "New State:"
  Text 20, 90, 90, 10, "Forwarding ADDR provided:"
  Text 35, 170, 65, 10, "Updated Address:"
  Text 85, 70, 35, 10, "Zip Code:"
  Text 10, 10, 95, 10, "Returned Mail Received on:"
  Text 15, 50, 35, 10, "From City:"
  Text 10, 130, 35, 10, "New City:"
  Text 35, 185, 70, 10, "County of new ADDR:"
  Text 10, 250, 90, 10, "Client have a MNsure case:"
  Text 20, 265, 75, 10, "MNsure case number:"
  Text 20, 280, 75, 10, "MNsure ADDR correct:"
EndDialog

01 Aitkin
02 Anoka
03 Becker
04 Beltrami
05 Benton
06 Big Stone
07 Blue Earth
08 Brown
09 Carlton
10 Carver
11 Cass
12 Chippewa
13 Chisago
14 Clay
15 Clearwater
16 Cook
17 Cottonwood
18 Crow Wing
19 Dakota
20 Dodge
21 Douglas
22 Faribault
23 Fillmore
24 Freeborn
25 Goodhue
26 Grant
27 Hennepin
28 Houston
29 Hubbard
30 Isanti
31 Itasca
32 Jackson
33 Kanabec
34 Kandiyohi
35 Kittson
36 Koochiching
37 Lac Qui Parle
38 Lake
39 Lake Of Woods
40 Le Sueur
41 Lincoln
42 Lyon
43 Mcleod
44 Mahnomen
45 Marshall
46 Martin
47 Meeker
48 Mille Lacs
49 Morrison
50 Mower
51 Murray
52 Nicollet
53 Nobles
54 Norman
55 Olmsted
56 Otter Tail
57 Pennington
58 Pine
59 Pipestone
60 Polk
61 Pope
62 Ramsey
63 Red Lake
64 Redwood
65 Renville
66 Rice
67 Rock
68 Roseau
69 St. Louis
70 Scott
71 Sherburne
72 Sibley
73 Stearns
74 Steele
75 Stevens
76 Swift
77 Todd
78 Traverse
79 Wabasha
80 Wadena
81 Waseca
82 Washington
83 Watonwan
84 Wilkin
85 Winona
86 Wright
87 Yellow Medicine
89 Out-of-State


01 Own Housing:
   Lease,
   Mortgage or
   Roommate
02 Family/Friends
   Due to
   Economic
   Hardship
03 Service
   Provider-
   Foster Care
   Group Home
   04 Hospital/
   Treatment/
   Detox/
   Nursing Home
05 Jail/Prison/
   Juvenile
   Detention
   Center
06 Hotel/Motel
07 Emergency
   Shelter
08 Place Not
   Meant for
   08 Place Not
   Meant for
   Housing
09 Declined
10 Unknown


BD Bois Forte -
   Deer Creek
BN Bois Forte -
   Nett Lake
BV Bois Forte -
   Vermillion Lk
FL Fond du Lac
GP Grand Portage
LL Leach Lake
LS Lower Sioux
ML Mille Lacs
PL Prairie Island
   Community
   RL Red Lake
SM Shakopee
   Mdewakanton
US Upper Sioux
WE White Earth

BeginDialog RETURNED_MAIL, 0, 0, 201, 280, "Mail has been returned with forwarding address in MN"
  Text 10, 15, 180, 35, maxis_addr
  CheckBox 10, 70, 70, 10, "Sent DHS-2919A", verifA_sent_checkbox
  CheckBox 85, 70, 65, 10, "Sent DHS-2952", SHEL_form_sent_checkbox
  CheckBox 10, 80, 65, 10, "Sent DHS-2402", CRF_sent_checkbox
  DropListBox 140, 105, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", update_ADDR
  EditBox 55, 125, 135, 15, new_ADDR
  EditBox 55, 145, 135, 15, new_CITY
  EditBox 55, 165, 35, 15, new_ZIP
  EditBox 165, 165, 25, 15, new_STATE
  DropListBox 55, 185, 35, 15, "Select One:"+chr(9)+"Aitkin"+chr(9)+"Anoka"+chr(9)+"Becker"+chr(9)+"Beltrami"+chr(9)+"Benton"+chr(9)+"Big Stone"+chr(9)+"Blue Earth"+chr(9)+"Brown"+chr(9)+"Carlton"+chr(9)+"Carver", county_code
  DropListBox 155, 185, 35, 15, "Select One:"+chr(9)+"Residence"+chr(9)+"Mailing"+chr(9)+"Both"+chr(9)+"Unknown", residence_addr
  DropListBox 55, 200, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", homeless_addr
  DropListBox 155, 200, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", living_situation
  DropListBox 55, 215, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", reservation_addr
  DropListBox 125, 215, 65, 15, "Select One:"+chr(9)+"Bois Forte - Deer Creek   "+chr(9)+"Bois Forte - Nett Lake    "+chr(9)+"Bois Forte - Vermillion Lk"+chr(9)+"Fond du Lac  ", reservation_name
  DropListBox 125, 240, 65, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", METS_ADDR
  EditBox 135, 255, 55, 15, MNsure_number
  EditBox 50, 275, 140, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 75, 275, 50, 15
    CancelButton 140, 275, 50, 15
  Text 35, 150, 20, 10, "City:"
  Text 140, 170, 20, 10, "State:"
  Text 50, 110, 85, 10, "Script to update address?"
  Text 10, 190, 30, 10, "County:"
  Text 20, 170, 35, 10, "Zip code:"
  GroupBox 5, 95, 190, 140, "New Address:"
  Text 30, 130, 20, 10, "Street:"
  Text 5, 275, 40, 10, "Other notes:"
  Text 100, 220, 25, 10, "Name:"
  GroupBox 5, 5, 190, 50, "Address in MAXIS:"
  Text 10, 220, 45, 10, "Reservation:"
  Text 5, 240, 95, 10, "METS correspondence sent:"
  Text 10, 205, 35, 10, "Homeless:"
  GroupBox 5, 55, 190, 40, "Verification Request Form"
  Text 100, 190, 40, 10, "Is address:"
  Text 100, 205, 55, 10, "Living situation:"
  Text 5, 260, 70, 10, "METS case number:"
EndDialog

Dim myBtn

myBtn = Dialog(RETURNED_MAIL)
MsgBox "The user pressed button " & myBtn





'starts the EVF received case note dialog
DO
	err_msg = ""
	'starts the Returned Mail dialog
	Dialog RETURNED_MAIL
	'asks if you want to cancel and if "yes" is selected sends StopScript
	cancel_confirmation
	'checks that there is a date in the date received box
	IF IsDate (date_received) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a date in mm/dd/yy for date received."
	'checks if the ADDR rec'd from is filled in
	IF from_ADDR = "" THEN err_msg = err_msg & vbCr & "You must enter the ADDR rec'd from."
	'checks if the from City has been entered
	IF from_CITY = "" THEN err_msg = err_msg & vbCr & "You must enter the City rec'd from."
	'checks if the from State has been entered
	IF from_STATE = "" THEN err_msg = err_msg & vbCr & "You must enter the State rec'd from."
	'checks if the from Zip has been entered
	IF from_ZIP = "" THEN err_msg = err_msg & vbCr & "You must enter the Zip rec'd from."
	'checks if a forwarding ADDR has been provided and new ADDR entered if provided
	IF forwarding_ADDR = "Yes" and new_ADDR = "" THEN err_msg = err_msg & vbCr & "You must input the new ADDR."
	'checks if a forwarding ADDR has been provided and new State entered if provided
	IF forwarding_ADDR = "Yes" and new_CITY = "" THEN err_msg = err_msg & vbCr & "You must input the new City."
	'checks if a forwarding ADDR has been provided and new State entered if provided
	IF forwarding_ADDR = "Yes" and new_STATE = "" THEN err_msg = err_msg & vbCr & "You must input the new State."
	'checks if STAT/ADDR was updated when a forwarding ADDR was provided
	IF forwarding_ADDR = "Yes" and updated_ADDR = "No" THEN err_msg = err_msg & vbCr & "You must update the address when a forwarding address is rec'd."
	'checks if a forwarding ADDR has been provided
	IF forwarding_ADDR = "Yes" and new_ZIP = "" THEN err_msg = err_msg & vbCr & "You must input the new Zip Code."
	'checks if a forwarding ADDR has been provided
	IF forwarding_ADDR = "Yes" and new_COUNTY = "" THEN err_msg = err_msg & vbCr & "You must input the County of the new ADDR."
	'checks if client is active on MNsure question has been answered
	IF MNsure_active = "Select" THEN err_msg = err_msg & vbCr & "You must select if the client has a MNsure case or not."
	'checks if MNsure case number has been entered on a MNsure active case
	IF MNsure_active = "Yes" and MNsure_number = "" THEN err_msg = err_msg & vbCr & "You must enter the MNsure case number."
	'checks if MNsure ADDR updated
	IF MNsure_active = "Yes" and MNsure_ADDR = "N/A" THEN err_msg = err_msg & vbCr & "You must select if the MNsure ADDR is correct."
	'checks if notes/actions taken were entered
	IF other_notes = "" THEN err_msg = err_msg & vbCr & "You must enter action taken/misc notes."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'assigns a value to the ADDR_status variable based on the value of the complete variable
IF forwarding_ADDR = "Yes" THEN ADDR_status = "a forwarding ADDR."
IF forwarding_ADDR = "No" THEN ADDR_status = "no forwarding ADDR."

'assigns a value to the MNsure variable based on the value of MNsure_active
IF MNsure_active = "Yes" THEN MNsure = "MNsure case"
IF MNsure_active = "No" THEN MNsure = "Non-MNsure"

'converts the old and new ADDR to all CAPS
from_ADDR = UCase(from_ADDR)
from_CITY = UCase(from_CITY)
from_STATE = UCase(from_STATE)
new_ADDR = UCase(new_ADDR)
new_CITY = UCase(new_CITY)
new_STATE = UCase(new_STATE)
new_COUNTY = UCase(new_COUNTY)

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
call check_for_MAXIS (false)

'starts a blank case note
call start_a_blank_case_note

'this enters the actual case note info
call write_variable_in_CASE_NOTE("***Returned Mail received " & date_received & " with " & ADDR_status & "*** " & MNsure)
call write_bullet_and_variable_in_CASE_NOTE("From ADDR", from_ADDR)
call write_variable_in_CASE_NOTE("             " & from_CITY & ", " & from_STATE & " " & from_ZIP)
IF forwarding_ADDR = "Yes" THEN call write_variable_in_CASE_NOTE("* Address updated to: " & new_ADDR)
IF forwarding_ADDR = "Yes" THEN call write_variable_in_CASE_NOTE("                      " & new_CITY & ", " & new_STATE & " " & new_Zip)
IF forwarding_ADDR = "Yes" THEN call write_variable_in_CASE_NOTE("* New ADDR is in " & new_COUNTY & " COUNTY.")
If verifA_sent_checkbox = 1 then call write_variable_in_CASE_NOTE("* Verification Request Form A sent. **Auto TIKL set**")
If SHEL_form_sent_checkbox = 1 then call write_variable_in_CASE_NOTE("* Shelter Verification Form sent.")
If CRF_sent_checkbox = 1 then call write_variable_in_CASE_NOTE("* Change Report Form sent.")
call write_bullet_and_variable_in_CASE_NOTE("Returned Mail resent", mail_resent)
If returned_mail_resent_list = "Yes" then call write_variable_in_CASE_NOTE("* Returned Mail resent to client.")
If returned_mail_resent_list = "No" then call write_variable_in_CASE_NOTE("* Returned Mail not resent to client.")
call write_variable_in_CASE_NOTE("* " & MNsure & " " & MNsure_number)
call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Checks if this is a MNsure case and pops up a message box with instructions if the ADDR is incorrect.
IF MNsure_active = "Yes" and MNsure_ADDR = "No" THEN MsgBox "Please update the MNsure ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. HPU Case Manitenance - Action Needed Log)"

'creates a message box reminding the worker to review their case note prior to Auto-TIKLing.
IF verifA_sent_checkbox = 1 THEN MsgBox "Please review your case note for accuracy. When you click OK or press enter the script will enter an Auto-TIKL for you."

'Checks if a DHS2919A mailed and sets a TIKL for the return of the info.
IF verifA_sent_checkbox = 1 THEN
	call navigate_to_MAXIS_screen("dail", "writ")

	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(date, 10, 5, 18)

	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("ADDR verification requested via 2919A after returned mail being rec'd should have returned by now. If not received, take appropriate action." )
	transmit
	PF3

	'Success message
	MsgBox "Success! TIKL has been sent for 10 days from now for the ADDR verification requested via 2919A."

End if

script_end_procedure("")
