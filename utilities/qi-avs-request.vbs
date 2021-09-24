name_of_script = "UTILITIES - QI AVS REQUEST.vbs"
start_time = timer
STATS_counter = 0                     	'sets the stats counter at zero
STATS_manualtime = 300               	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "M"       		'C is for each CASE
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
call changelog_update("09/24/2021", "GitHub Issue #583 Updates made to ensure email has information went sent to QI", "MiKayla Handley, Hennepin County")
call changelog_update("09/08/2021", "Added date completed AVS form rec'd to dialog and reminder that completed AVS needs to be on file prior to submitting AVS request.", "Ilse Ferris, Hennepin County")
call changelog_update("09/30/2020", "Updated closing message.", "Ilse Ferris, Hennepin County")
call changelog_update("03/10/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
function read_ADDR_panel(addr_eff_date, line_one, line_two, city, state, zip, county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, phone_one, type_one, phone_two, type_two, phone_three, type_three, updated_date)
    Call navigate_to_MAXIS_screen("STAT", "ADDR")
    EMReadScreen line_one, 22, 6, 43
    EMReadScreen line_two, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 7, 9, 43
    EMReadScreen county_line, 2, 9, 66
    EMReadScreen verif_line, 2, 9, 74
    EMReadScreen homeless, 1, 10, 43
    EMReadScreen ind_reservation, 1, 10, 74
    EMReadScreen living_sit, 2, 11, 43
	EMReadScreen res_name, 2, 11, 74
    line_one = replace(line_one, "_", "")
    line_two = replace(line_two, "_", "")
    city = replace(city_line, "_", "")
    state = state_line
    zip = replace(zip_line, "_", "")

    If county_line = "01" Then county = "01 - Aitkin"
    If county_line = "02" Then county = "02 - Anoka"
    If county_line = "03" Then county = "03 - Becker"
    If county_line = "04" Then county = "04 - Beltrami"
    If county_line = "05" Then county = "05 - Benton"
    If county_line = "06" Then county = "06 - Big Stone"
    If county_line = "07" Then county = "07 - Blue Earth"
    If county_line = "08" Then county = "08 - Brown"
    If county_line = "09" Then county = "09 - Carlton"
    If county_line = "10" Then county = "10 - Carver"
    If county_line = "11" Then county = "11 - Cass"
    If county_line = "12" Then county = "12 - Chippewa"
    If county_line = "13" Then county = "13 - Chisago"
    If county_line = "14" Then county = "14 - Clay"
    If county_line = "15" Then county = "15 - Clearwater"
    If county_line = "16" Then county = "16 - Cook"
    If county_line = "17" Then county = "17 - Cottonwood"
    If county_line = "18" Then county = "18 - Crow Wing"
    If county_line = "19" Then county = "19 - Dakota"
    If county_line = "20" Then county = "20 - Dodge"
    If county_line = "21" Then county = "21 - Douglas"
    If county_line = "22" Then county = "22 - Faribault"
    If county_line = "23" Then county = "23 - Fillmore"
    If county_line = "24" Then county = "24 - Freeborn"
    If county_line = "25" Then county = "25 - Goodhue"
    If county_line = "26" Then county = "26 - Grant"
    If county_line = "27" Then county = "27 - Hennepin"
    If county_line = "28" Then county = "28 - Houston"
    If county_line = "29" Then county = "29 - Hubbard"
    If county_line = "30" Then county = "30 - Isanti"
    If county_line = "31" Then county = "31 - Itasca"
    If county_line = "32" Then county = "32 - Jackson"
    If county_line = "33" Then county = "33 - Kanabec"
    If county_line = "34" Then county = "34 - Kandiyohi"
    If county_line = "35" Then county = "35 - Kittson"
    If county_line = "36" Then county = "36 - Koochiching"
    If county_line = "37" Then county = "37 - Lac Qui Parle"
    If county_line = "38" Then county = "38 - Lake"
    If county_line = "39" Then county = "39 - Lake Of Woods"
    If county_line = "40" Then county = "40 - Le Sueur"
    If county_line = "41" Then county = "41 - Lincoln"
    If county_line = "42" Then county = "42 - Lyon"
    If county_line = "43" Then county = "43 - Mcleod"
    If county_line = "44" Then county = "44 - Mahnomen"
    If county_line = "45" Then county = "45 - Marshall"
    If county_line = "46" Then county = "46 - Martin"
    If county_line = "47" Then county = "47 - Meeker"
    If county_line = "48" Then county = "48 - Mille Lacs"
    If county_line = "49" Then county = "49 - Morrison"
    If county_line = "50" Then county = "50 - Mower"
    If county_line = "51" Then county = "51 - Murray"
    If county_line = "52" Then county = "52 - Nicollet"
    If county_line = "53" Then county = "53 - Nobles"
    If county_line = "54" Then county = "54 - Norman"
    If county_line = "55" Then county = "55 - Olmsted"
    If county_line = "56" Then county = "56 - Otter Tail"
    If county_line = "57" Then county = "57 - Pennington"
    If county_line = "58" Then county = "58 - Pine"
    If county_line = "59" Then county = "59 - Pipestone"
    If county_line = "60" Then county = "60 - Polk"
    If county_line = "61" Then county = "61 - Pope"
    If county_line = "62" Then county = "62 - Ramsey"
    If county_line = "63" Then county = "63 - Red Lake"
    If county_line = "64" Then county = "64 - Redwood"
    If county_line = "65" Then county = "65 - Renville"
    If county_line = "66" Then county = "66 - Rice"
    If county_line = "67" Then county = "67 - Rock"
    If county_line = "68" Then county = "68 - Roseau"
    If county_line = "69" Then county = "69 - St. Louis"
    If county_line = "70" Then county = "70 - Scott"
    If county_line = "71" Then county = "71 - Sherburne"
    If county_line = "72" Then county = "72 - Sibley"
    If county_line = "73" Then county = "73 - Stearns"
    If county_line = "74" Then county = "74 - Steele"
    If county_line = "75" Then county = "75 - Stevens"
    If county_line = "76" Then county = "76 - Swift"
    If county_line = "77" Then county = "77 - Todd"
    If county_line = "78" Then county = "78 - Traverse"
    If county_line = "79" Then county = "79 - Wabasha"
    If county_line = "80" Then county = "80 - Wadena"
    If county_line = "81" Then county = "81 - Waseca"
    If county_line = "82" Then county = "82 - Washington"
    If county_line = "83" Then county = "83 - Watonwan"
    If county_line = "84" Then county = "84 - Wilkin"
    If county_line = "85" Then county = "85 - Winona"
    If county_line = "86" Then county = "86 - Wright"
    If county_line = "87" Then county = "87 - Yellow Medicine"
    If county_line = "89" Then county = "89 - Out-of-State"

    If homeless = "Y" Then homeless = "Yes"
    If homeless = "N" Then homeless = "No"
    If ind_reservation = "Y" Then ind_reservation = "Yes"
    If ind_reservation = "N" Then ind_reservation = "No"

    If verif_line = "SF" Then verif = "SF - Shelter Form"
    If verif_line = "Co" Then verif = "CO - Coltrl Stmt"
    If verif_line = "MO" Then verif = "MO - Mortgage Papers"
    If verif_line = "TX" Then verif = "TX - Prop Tax Stmt"
    If verif_line = "CD" Then verif = "CD - Contrct for Deed"
    If verif_line = "UT" Then verif = "UT - Utility Stmt"
    If verif_line = "DL" Then verif = "DL - Driver Lic/State ID"
    If verif_line = "OT" Then verif = "OT - Other Document"
    If verif_line = "NO" Then verif = "NO - No Ver Prvd"
    If verif_line = "?_" Then verif = "? - Delayed"
    If verif_line = "__" Then verif = "Blank"

    If living_sit = "__" Then living_sit = "Blank"
    If living_sit = "01" Then living_sit = "01 - Own Housing (lease, mortgage, or roomate)"
    If living_sit = "02" Then living_sit = "02 - Family/Friends due to economic hardship"
    If living_sit = "03" Then living_sit = "03 - Servc prvdr- foster/group home"
    If living_sit = "04" Then living_sit = "04 - Hospital/Treatment/Detox/Nursing Home"
    If living_sit = "05" Then living_sit = "05 - Jail/Prison//Juvenile Det."
    If living_sit = "06" Then living_sit = "06 - Hotel/Motel"
    If living_sit = "07" Then living_sit = "07 - Emergency Shelter"
    If living_sit = "08" Then living_sit = "08 - Place not meant for Housing"
    If living_sit = "09" Then living_sit = "09 - Declined"
    If living_sit = "10" Then living_sit = "10 - Unknown"

	If res_name = "__" Then res_name = "Blank"
	If res_name = "BD" Then res_name = "Bois Forte - Deer Creek"
	If res_name = "BN" Then res_name = "Bois Forte - Nett Lake"
	If res_name = "BV" Then res_name = "Bois Forte - Vermillion Lk"
	If res_name = "FL" Then res_name = "Fond du Lac"
	If res_name = "GP" Then res_name = "Grand Portage"
	If res_name = "LL" Then res_name = "Leach Lake"
	If res_name = "LS" Then res_name = "Lower Sioux"
	If res_name = "ML" Then res_name = "Mille Lacs"
	If res_name = "PL" Then res_name = "Prairie Island Community"
	If res_name = "RL" Then res_name = "Red Lake"
	If res_name = "SM" Then res_name = "Shakopee Mdewakanton"
	If res_name = "US" Then res_name = "Upper Sioux"
	If res_name = "WE" Then res_name = "White Earth"

    EMReadScreen addr_eff_date, 8, 4, 43
    EMReadScreen addr_future_date, 8, 4, 66
    EMReadScreen mail_line_one, 22, 13, 43
    EMReadScreen mail_line_two, 22, 14, 43
    EMReadScreen mail_city, 15, 15, 43
    EMReadScreen mail_state, 2, 16, 43
    EMReadScreen mail_zip, 7, 16, 52

    addr_eff_date = replace(addr_eff_date, " ", "/")
    addr_future_date = trim(addr_future_date)
    addr_future_date = replace(addr_future_date, " ", "/")
    mail_line_one = replace(mail_line_one, "_", "")
    mail_line_two = replace(mail_line_two, "_", "")
    mail_city = replace(mail_city, "_", "")
    mail_state = replace(mail_state, "_", "")
    mail_zip = replace(mail_zip, "_", "")

	EMReadScreen phone_one, 14, 17, 45
	EMReadScreen phone_two, 14, 18, 45
	EMReadScreen phone_three, 14, 19, 45
	EMReadScreen type_one, 1, 17, 67
	EMReadScreen type_two, 1, 18, 67
	EMReadScreen type_three, 1, 19, 67

	phone_one = "(" & replace(replace(replace(phone_one, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_one = "(___) ___ - ____" Then phone_one = ""
	If type_one = "_" Then type_one = "Unknown"
	If type_one = "H" Then type_one = "Home"
	If type_one = "W" Then type_one = "Work"
	If type_one = "C" Then type_one = "Cell"
	If type_one = "M" Then type_one = "Message"
	If type_one = "T" Then type_one = "TTY/TDD"

	phone_two = "(" & replace(replace(replace(phone_two, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_two = "(___) ___ - ____" Then phone_two = ""
	If type_two = "_" Then type_two = "Unknown"
	If type_two = "H" Then type_two = "Home"
	If type_two = "W" Then type_two = "Work"
	If type_two = "C" Then type_two = "Cell"
	If type_two = "M" Then type_two = "Message"
	If type_two = "T" Then type_two = "TTY/TDD"

	phone_three = "(" & replace(replace(replace(phone_three, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_three = "(___) ___ - ____" Then phone_three = ""
	If type_three = "_" Then type_three = "Unknown"
	If type_three = "H" Then type_three = "Home"
	If type_three = "W" Then type_three = "Work"
	If type_three = "C" Then type_three = "Cell"
	If type_three = "M" Then type_three = "Message"
	If type_three = "T" Then type_three = "TTY/TDD"

	EMReadScreen updated_date, 8, 21, 55
	updated_date = replace(updated_date, " ", "/")
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
'Connecting to BlueZone
EMConnect ""
'Grabs the case number
CALL MAXIS_case_number_finder (MAXIS_case_number)
closing_message = "Request for Account Validation Service (AVS) email has been sent." 'setting up closing_message or possible additions later based on conditions
'----------------------------------------------------------------------------------------------------Initial dialog
appl_type = "Application"
applicant_type = "Applicant"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 211, 160, "AVS Request"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 185, 5, 20, 15, HH_size
  EditBox 155, 25, 50, 15, avs_form_date
  DropListBox 80, 60, 125, 15, "Select One:"+chr(9)+"Applicant"+chr(9)+"Spouse", applicant_type
  DropListBox 80, 80, 125, 15, "Select One:"+chr(9)+"Application"+chr(9)+"Renewal", appl_type
  DropListBox 80, 100, 125, 15, "Select One:"+chr(9)+"BI-Brain Injury Waiver"+chr(9)+"BX-Blind"+chr(9)+"CA-Community Alt. Care"+chr(9)+"DD-Developmental Disa Waiver"+chr(9)+"DP-MA for Employed Pers w/ Disa"+chr(9)+"DX-Disability"+chr(9)+"EH-Emergency Medical Assistance"+chr(9)+"EW-Elderly Waiver"+chr(9)+"EX-65 and Older"+chr(9)+"LC-Long Term Care"+chr(9)+"MP-QMB SLMB Only"+chr(9)+"QI-QI"+chr(9)+"QW-QWD", MA_type
  DropListBox 80, 120, 125, 15, "Select One:"+chr(9)+"NA-No Spouse"+chr(9)+"YES"+chr(9)+"NO", spouse_deeming
  ButtonGroup ButtonPressed
    OkButton 110, 140, 45, 15
    CancelButton 160, 140, 45, 15
  Text 5, 65, 50, 10, "Applicant Type:"
  Text 5, 85, 65, 10, "Application Type:"
  Text 5, 105, 55, 10, "Request Type:"
  Text 5, 125, 35, 10, "Deeming:"
  Text 150, 10, 30, 10, "HH Size:"
  Text 5, 30, 120, 10, "Date Completed AVS Form Received:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 45, 200, 10, "AVS form must be complete and valid to submit AVS Request."
EndDialog

DO
    DO
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If HH_size = "" or IsNumeric(HH_size) = False then err_msg = err_msg & vbNewLine & "* Please enter a valid household composition size."
        If trim(avs_form_date) = "" or isdate(avs_form_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date the completed AVS form was received in the agency."
		IF applicant_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the applicant type."
		IF appl_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the application type."
		IF MA_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the MA request type."
		IF spouse_deeming = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select if the spouse is deeming."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password_without_transmit(are_we_passworded_out)
Loop until are_we_passworded_out = false
CALL check_for_MAXIS(False)
CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv) 'navigating to stat prog to gather the application information
IF is_this_priv = TRUE THEN script_end_procedure("PRIV case, cannot access/update. The script will now end.")

EMReadScreen application_date, 8, 12, 33 'Reading the HC app date from PROG
application_date = replace(application_date, " ", "/")
IF application_date = "__/__/__"  THEN script_end_procedure("*** No application date ***" & vbNewLine & "Need to have pending or active HC care to request AVS.")

CALL HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.

CALL navigate_to_MAXIS_screen("STAT", "MEMB") 'navigating to stat memb to gather the ref number and name.

DO
    CALL HH_member_custom_dialog(HH_member_array)
    IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
LOOP UNTIL uBound(HH_member_array) <> -1

CALL get_county_code
EMReadscreen current_county, 4, 21, 21
If lcase(current_county) <> worker_county_code THEN script_end_procedure("Out of County case, cannot access/update. The script will now end.")

'Establishing array
avs_membs = 0       'incrementor for array
DIM avs_members_array()  'Declaring the array this is what this list is
ReDim avs_members_array(phone_type_three_const, 0)  'Resizing the array 'redimmed to the size of the last constant  'that ,list is going to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const maxis_case_number_const  	 	= 0 '=  Maxis'
const member_number_const   		= 1 '=  Member Number
const client_first_name_const       = 2 '=  First Name MEMB
const client_last_name_const        = 3 '=  Last Name MEMB
const client_mid_name_const    	    = 4 '=  Middle initial MEMB
const client_DOB_const   		    = 5 '=  Date of Birth MEMB
const client_ssn_const		        = 6 '=  SSN
const client_age_const	            = 7 '=  age MEMB
const client_sex_const			    = 8 '=  client sex
const addr_eff_date_const	  		= 9 '=	addr_eff_date
const resi_line_one_const	  		= 10'= 	resi_line_one
const resi_line_two_const	   		= 11 '= resi_line_two
const resi_city_const				= 12'= 	resi_city
const resi_state_const	     		= 13'= 	resi_state
const resi_zip_const     			= 14'= 	resi_zip
const resi_county_const     		= 15'= 	resi_county
const verif_const	  				= 16'= 	verif
const homeless_const     			= 17 '= homeless
const ind_reservation_const  		= 18 '= ind_reservation
const living_sit_const	     		= 19 '= living_sit
const res_name_const    			= 20'= 	res_name
const mail_line_one_const    		= 21'= 	mail_line_one
const mail_line_two_const     		= 22'=  mail_line_two
const mail_city_const   			= 23'= 	mail_city
const mail_state_const    			= 24'= 	mail_state
const mail_zip_const    			= 25'= 	mail_zip
const phone_numb_one_const    		= 26'= 	phone_numb_one
const phone_type_one_const    		= 27'= 	phone_type_one
const phone_numb_two_const     		= 28'= 	phone_numb_two
const phone_type_two_const  	   	= 29'= 	phone_type_two
const phone_numb_three_const     	= 30'= 	phone_numb_three
const phone_type_three_const    	= 31'= 	phone_type_three

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
FOR EACH person IN HH_member_array
    CALL write_value_and_transmit(person, 20, 76) 'reads the reference number, last name, first name, and THEN puts it into an array YOU HAVENT defined the avs_members_array yet
    EMReadscreen ref_nbr, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen MEMB_number, 3, 4, 33
    EMReadscreen last_name, 25, 6, 30
    EMReadscreen first_name, 12, 6, 63
    EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
    EMReadscreen client_SSN, 11, 7, 42
    If client_ssn = "___ __ ____" then client_ssn = ""
    last_name = trim(replace(last_name, "_", "")) & " "
    first_name = trim(replace(first_name, "_", "")) & " "
    mid_initial = replace(mid_initial, "_", "")
    EMReadScreen client_age, 2, 8, 76
    IF client_age = "  " THEN client_age = 0
    client_age = client_age * 1
	EMReadScreen client_sex, 1, 9, 42
    ReDim Preserve avs_members_array(phone_type_three_const, avs_membs)  'redimmed to the size of the last constant
    avs_members_array(member_number_const,     avs_membs) = ref_nbr
    avs_members_array(client_first_name_const, avs_membs) = first_name
    avs_members_array(client_last_name_const,  avs_membs) = last_name
    avs_members_array(client_mid_name_const,   avs_membs) = mid_initial
    avs_members_array(client_DOB_const,        avs_membs) = client_DOB
    avs_members_array(client_ssn_const,        avs_membs) = client_SSN
    avs_members_array(client_age_const,        avs_membs) = client_age
	avs_members_array(client_sex_const,        avs_membs) = client_sex
    avs_membs = avs_membs + 1 ' can only be used because we havent reset or redefined this incrementor'
	STATS_counter = STATS_counter + 1
NEXT

CALL navigate_to_MAXIS_screen("STAT", "MEMI")
EMReadScreen marital_status, 1, 7, 40
EMReadScreen spouse_ref_nbr, 02, 09, 49
spouse_ref_nbr = replace(spouse_ref_nbr, "_", "")
IF marital_status = "M" and spouse_ref_nbr <> "" THEN client_married = TRUE
IF spouse_deeming = "YES" and spouse_ref_nbr = "" THEN
	BeginDialog Dialog1, 0, 0, 176, 160, "Spouse not found on MEMB"
      EditBox 55, 5, 115, 15, spouse_first_name
      EditBox 55, 25, 115, 15, spouse_last_name
      EditBox 55, 45, 35, 15, spouse_mid_name
      EditBox 120, 45, 50, 15, spouse_SSN_number
      EditBox 55, 65, 55, 15, spouse_DOB
      EditBox 150, 65, 20, 15, spouse_age
      DropListBox 55, 85, 55, 15, "Select One:"+chr(9)+"Female"+chr(9)+"Male"+chr(9)+"Unknown"+chr(9)+"Undetermined", spouse_gender_dropdown
      EditBox 5, 120, 165, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 75, 140, 45, 15
        CancelButton 125, 140, 45, 15
      Text 5, 10, 40, 10, "First Name:"
      Text 5, 30, 40, 10, "Last Name: "
      Text 5, 50, 50, 10, " Middle Initial: "
      Text 100, 50, 20, 10, "SSN: "
      Text 5, 70, 45, 10, "Date of Birth: "
      Text 130, 70, 15, 10, "Age: "
      Text 5, 105, 160, 10, "Please explain why they are not listed in maxis: "
      Text 5, 90, 30, 10, "Gender: "
    EndDialog

	DO
	 	DO
	 		err_msg = ""
	 		Dialog Dialog1
	 		cancel_without_confirmation
	 		If spouse_first_name = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's first name."
	        If spouse_last_name = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's last name."
	        If spouse_SSN_number = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's social security number."
	        If spouse_DOB = "" then err_msg = err_msg & vbNewLine & "Please enter the spouse's date of birth."
	        If spouse_gender_dropdown = "Select One:" then err_msg = err_msg & vbNewLine & "Please select the spouse's gender."
	        If other_notes = "" then err_msg = err_msg & vbNewLine & "Please enter the reason this client is not listed in MAXIS."
	 		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	 	LOOP UNTIL err_msg = ""
	 	CALL check_for_password_without_transmit(are_we_passworded_out)
	Loop until are_we_passworded_out = false

	ReDim Preserve avs_members_array(phone_type_three_const, avs_membs)  'redimmed to the size of the last constant
    avs_members_array(member_number_const,     avs_membs) = spouse_ref_nbr
    avs_members_array(client_first_name_const, avs_membs) = spouse_first_name
    avs_members_array(client_last_name_const,  avs_membs) = spouse_last_name
    avs_members_array(client_mid_name_const,   avs_membs) = spouse_mid_name
    avs_members_array(client_DOB_const,        avs_membs) = spouse_DOB
    avs_members_array(client_ssn_const,        avs_membs) = spouse_SSN_number
    avs_members_array(client_age_const,        avs_membs) = spouse_age
	avs_members_array(client_sex_const,        avs_membs) = spouse_gender_dropdown
	client_married = TRUE
END IF
CALL read_ADDR_panel(addr_eff_date, line_one, line_two, city, state, zip, county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, phone_one, type_one, phone_two, type_two, phone_three, type_three, updated_date)

team_email = "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us"

FOR avs_membs = 0 to Ubound(avs_members_array, 2) 'start at the zero person and go to each of the selected people '
    member_info = member_info & "A signed AVS form was received for Member # " & avs_members_array(member_number_const, avs_membs) & vbNewLine & avs_members_array(client_first_name_const, avs_membs) & " " & avs_members_array(client_mid_name_const, avs_membs) & " " & avs_members_array(client_last_name_const, avs_membs)  &  vbCr & "DOB: " & avs_members_array(client_DOB_const,  avs_membs) & vbcr & "SSN of Resident: " & avs_members_array(client_ssn_const,  avs_membs) & vbcr & "Gender: " & avs_members_array(client_sex_const, avs_membs)
	member_info = member_info & vbNewLine & "AVS Form Received Date: " & avs_form_date & vbcr & "MA type: " & MA_type & vbcr & "HH size: " & HH_size & vbcr & "Applicant Type: " & applicant_type & vbcr & "Application Type: " & appl_type & vbNewLine & "Residential Address: " & vbNewLine & line_one & " " & line_two & vbcr & city & ", " & state & " " & zip
	If trim(mail_line_one) <> "" THEN member_info = member_info & "Mailing address: " & mail_line_one & vbcr & mail_line_two & vbcr & mail_city & vbcr & mail_state & vbcr & mail_zip & vbcr & phone_one & " Phone: " & type_one & " - " & phone_two & " - " & phone_three
	IF client_married = TRUE THEN member_info = member_info & vbNewLine & "Spouse: " & spouse_deeming & vbcr & "Spouse Member # " & avs_members_array(member_number_const, avs_membs) & vbcr & "Spouse First Name: " & avs_members_array(client_first_name_const, avs_membs) & vbcr & "Spouse Last Name: " & avs_members_array(client_last_name_const, avs_membs) & vbcr & "Spouse Social Security Number: " & avs_members_array(client_ssn_const,  avs_membs) & vbcr & "Spouse Gender: " & avs_members_array(client_sex_const, avs_membs) & vbcr & "Spouse Date of birth: " & avs_members_array(client_DOB_const, avs_membs) & " " & other_notes
NEXT

CALL find_user_name(the_person_running_the_script)' this is for the signature in the email'

'Creating the email
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachmentsend_email)
Call create_outlook_email(team_email, "", "AVS initial run requests case #" & MAXIS_case_number, member_info & vbNewLine & vbNewLine & "Submitted By: " & vbNewLine & the_person_running_the_script, "", FALSE)   'will create email, will send.

script_end_procedure_with_error_report(closing_message)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/30/2021
'--Tab orders reviewed & confirmed----------------------------------------------08/30/2021
'--Mandatory fields all present & Reviewed--------------------------------------08/30/2021
'--All variables in dialog match mandatory fields-------------------------------08/30/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/09/21
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------08/30/2021
'--Out-of-County handling reviewed----------------------------------------------08/30/2021
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/09/21
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------09/09/21
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------08/30/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------08/30/2021
'--Comment Code-----------------------------------------------------------------09/09/21
'--Update Changelog for release/update------------------------------------------09/09/21
'--Remove testing message boxes-------------------------------------------------09/09/21
'--Remove testing code/unnecessary code-----------------------------------------09/09/21
'--Review/update SharePoint instructions----------------------------------------09/09/21
'--Review Best Practices using BZS page ----------------------------------------09/09/21
'--Review script information on SharePoint BZ Script List-----------------------09/09/21
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/09/21
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/09/21
'--Complete misc. documentation (if applicable)---------------------------------09/09/21
'--Update project team/issue contact (if applicable)----------------------------09/09/21
